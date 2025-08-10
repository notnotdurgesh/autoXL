import { GoogleGenerativeAI, FunctionCallingMode } from '@google/generative-ai';
import type { Tool, GenerativeModel, ChatSession } from '@google/generative-ai';
import { allFunctionDeclarations } from './functionDeclarations';
import type { ChatMessage, FileAttachment } from '../../types/ai.types';
import { getFriendlyCompletionMessage } from '../../utils/minimalFunctionMappings';
import { systemInstruction } from '../../utils/prompts';

export class GeminiService {
  private genAI: GoogleGenerativeAI;
  private model!: GenerativeModel;
  private chat!: ChatSession;
  private functionHandlers: Map<string, (args: Record<string, unknown>) => Promise<unknown>>;
  private modelId: string = 'gemini-2.5-flash-lite';
  private apiKey: string;

  constructor(apiKey: string) {
    this.apiKey = apiKey;
    this.genAI = new GoogleGenerativeAI(apiKey);
    this.functionHandlers = new Map();
    this.initializeModel();
  }
  public setModel(modelId: string) {
    this.modelId = modelId;
    this.initializeModel();
  }

  // Best-effort image mime type inference from filename or data URL
  private inferImageMimeType(filenameOrName: string, dataUrl?: string): string {
    const lower = (filenameOrName || '').toLowerCase();
    if (lower.endsWith('.png')) return 'image/png';
    if (lower.endsWith('.jpg') || lower.endsWith('.jpeg')) return 'image/jpeg';
    if (lower.endsWith('.webp')) return 'image/webp';
    if (lower.endsWith('.gif')) return 'image/gif';
    if (lower.endsWith('.bmp')) return 'image/bmp';
    if (dataUrl && dataUrl.startsWith('data:')) {
      const semi = dataUrl.indexOf(';');
      const colon = dataUrl.indexOf(':');
      if (semi > colon && colon >= 0) return dataUrl.substring(colon + 1, semi);
    }
    return 'image/png';
  }

  private initializeModel() {
    // Create tools configuration with all function declarations
    const tools: Tool[] = [
      {
        functionDeclarations: allFunctionDeclarations,
      },
    ];

    // Initialize the model with function calling capabilities
    // Use a local variable typed as any to allow newer fields like `thinking` without TS friction
    // Use a lax type to accommodate optional fields not yet in SDK typings
    const generationConfig: Record<string, unknown> = {
      temperature: 0.5,  // Higher for creative problem solving
      topP: 0.9,        // Wider range for complex operations
      topK: 20,         // More options for multi-step planning
      maxOutputTokens: 2048,  // Allow longer answers with brief plan + citations
      candidateCount: 1,
    };

    this.model = this.genAI.getGenerativeModel({
      model: this.modelId,
      systemInstruction,
      tools,
      toolConfig: {
        functionCallingConfig: {
          mode: FunctionCallingMode.AUTO,  // Let AI decide when to use functions
        },
      },
      generationConfig,
    });

    // Start a chat session
    this.chat = this.model.startChat({
      history: [],
    });
  }

  // Register function handlers for spreadsheet operations
  public registerFunctionHandler(name: string, handler: (args: Record<string, unknown>) => Promise<unknown>) {
    this.functionHandlers.set(name, handler);
  }

  // Intent and safety helpers
  private isGreeting(message: string): boolean {
    const m = (message || '').trim().toLowerCase();
    return /^(hi|hello|hey|yo|sup|heya|hiya|hola|bonjour|namaste|salam)[!.\s]*$/.test(m);
  }
  // Send a message and handle function calls
  public async sendMessage(message: string, history: ChatMessage[] = [], attachments: FileAttachment[] = []): Promise<ChatMessage[]> {
    const responses: ChatMessage[] = [];

    try {
      // Convert history to Gemini format
      const formattedHistory = this.formatHistory(history);
      
      // Start a new chat with history (including prior image parts if any)
      this.chat = this.model.startChat({
        history: formattedHistory,
      });

      // Build user parts: text + any inline image data
      const userParts: Array<Record<string, unknown>> = [];
      if (message && message.trim().length > 0) {
        userParts.push({ text: message });
      }
      attachments
        .filter(att => att && typeof att.content === 'string' && att.type === 'image')
        .forEach(att => {
          const dataUrl = att.content as string;
          const commaIdx = dataUrl.indexOf(',');
          const base64Data = commaIdx >= 0 ? dataUrl.substring(commaIdx + 1) : dataUrl;
          userParts.push({
            inlineData: {
              mimeType: this.inferImageMimeType(att.name, dataUrl),
              data: base64Data,
            },
          });
        });

      // Safety short-circuit: if greeting, don't call tools
      if (this.isGreeting(message)) {
        responses.push({
          id: this.generateId(),
          role: 'assistant',
          content: 'Plan:\n- Greet the user\n- Await instruction\n\nAnswer:\nHi! What would you like to do with the sheet (analyze, format, add data, reorganize, or something else)?',
          timestamp: new Date(),
        });
        return responses;
      }

      // Send the message as multimodal parts (text + images)
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const result = await this.chat.sendMessage((userParts.length > 0 ? userParts : [{ text: message }]) as any);
      const response = result.response;

      // Check if there are function calls
      const functionCalls = response.functionCalls();
      
      if (functionCalls && functionCalls.length > 0) {
        // Log parallel execution if multiple functions
        if (functionCalls.length > 1) {
          console.log(`🚀 Executing ${functionCalls.length} functions in parallel`);
        }
        
        // Execute all functions in parallel
        const functionPromises = functionCalls.map(async (functionCall) => {
          const handler = this.functionHandlers.get(functionCall.name);
          
          if (handler) {
            try {
              // Execute the function
              const functionResult = await handler(functionCall.args as Record<string, unknown>);
              
              return {
                success: true,
                name: functionCall.name,
                result: functionResult,
              };
            } catch (error) {
              console.error(`Error in ${functionCall.name}:`, error);
              return {
                success: false,
                name: functionCall.name,
                error: String(error),
              };
            }
          } else {
            console.error(`No handler registered for function: ${functionCall.name}`);
            return {
              success: false,
              name: functionCall.name,
              error: 'Handler not found',
            };
          }
        });
        
        // Wait for all functions to complete
        const results = await Promise.all(functionPromises);
        
        // Create response messages for all function results
        const functionResponseParts: unknown[] = [];
        // Collect image parts to send in a separate message (cannot mix with FunctionResponse)
        const imagePartsToSend: unknown[] = [];
        
        results.forEach((result) => {
          if (result.success) {
            responses.push({
              id: this.generateId(),
              role: 'function',
              content: getFriendlyCompletionMessage(result.name),
              timestamp: new Date(),
              functionResponse: {
                name: result.name,
                result: result.result,
              },
            });
            
            functionResponseParts.push({
              functionResponse: {
                name: result.name,
                response: result.result,
              },
            });
            // Queue image for a separate message if scan_sheet produced one
            if (
              result.name === 'scan_sheet' &&
              result.result &&
              typeof result.result === 'object' &&
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              (result.result as any).imageDataUrl &&
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              typeof (result.result as any).imageDataUrl === 'string'
            ) {
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              const dataUrl = (result.result as any).imageDataUrl as string;
              const commaIdx = dataUrl.indexOf(',');
              const base64Data = commaIdx >= 0 ? dataUrl.substring(commaIdx + 1) : dataUrl;
              imagePartsToSend.push({
                inlineData: {
                  mimeType: 'image/png',
                  data: base64Data,
                },
              });
            }
          } else {
            responses.push({
              id: this.generateId(),
              role: 'function',
              content: `Couldn't complete the action`,
              timestamp: new Date(),
              functionResponse: {
                name: result.name,
                result: null,
                error: result.error,
              },
            });
          }
        });
        
        // Send all function results back to model at once
        if (functionResponseParts.length > 0) {
          try {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            let followUpResult = await this.chat.sendMessage(functionResponseParts as any);
            let followUpResponse = followUpResult.response;
            // If we have images to attach, send them in a separate message immediately after
            if (imagePartsToSend.length > 0) {
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              followUpResult = await this.chat.sendMessage(imagePartsToSend as any);
              followUpResponse = followUpResult.response;
            }
            
            // Check for additional function calls (compositional)
            const additionalCalls = followUpResponse.functionCalls();
            if (additionalCalls && additionalCalls.length > 0) {
              console.log(`🔄 Chaining ${additionalCalls.length} additional functions`);
              // Recursively handle additional function calls
              const additionalMessages = await this.handleFunctionCalls(additionalCalls);
              responses.push(...additionalMessages);
            } else {
              // No more function calls, add the text response (with optional citations)
              const followUpText = followUpResponse.text();
              if (followUpText) {
                const contentWithSources = this.appendGroundingCitationsSafely(followUpResponse, followUpText);
                responses.push({
                  id: this.generateId(),
                  role: 'assistant',
                  content: contentWithSources,
                  timestamp: new Date(),
                });
              }
            }
          } catch (error) {
            console.error('Error sending function responses:', error);
            responses.push({
              id: this.generateId(),
              role: 'assistant',
              content: `I completed the operations but encountered an issue with the response.`,
              timestamp: new Date(),
            });
          }
        }
      } else {
        // No function calls - check if this is an action request
        const text = response.text();
        
        // // Check if the response claims to have done something without calling functions
        // const actionWords = ['applied', 'deleted', 'formatted', 'added', 'created', 'set', 'changed', 'modified', 'executed', 'done', 'completed', 'will now', 'have now'];
        // const lowerText = text?.toLowerCase() || '';
        // const claimsAction = actionWords.some(word => lowerText.includes(word));
        
        // // Check if user message implies an action request
        // const userRequestsAction = /\b(do|apply|delete|format|add|create|set|change|modify|make|clear|remove|insert)\b/i.test(message);
        
        // if ((userRequestsAction || claimsAction) && !functionCalls && claimsAction) {
        //   // User requested an action OR AI claims to have done something without functions - that's wrong
        //   console.error('❌ AI claimed action without function calls!');
        //   responses.push({
        //     id: this.generateId(),
        //     role: 'assistant',
        //     content: '⚠️ I need to use actual functions to perform this action. Please try your request again.',
        //     timestamp: new Date(),
        //   });
        // } else 
        if (text) {
          // Legitimate text-only response (like answering questions)
          const contentWithSources = this.appendGroundingCitationsSafely(response, text);
          responses.push({
            id: this.generateId(),
            role: 'assistant',
            content: contentWithSources,
            timestamp: new Date(),
          });
        }
      }
    } catch (error) {
      console.error('Error sending message to Gemini:', error);
      responses.push({
        id: this.generateId(),
        role: 'assistant',
        content: `I encountered an error: ${error}. Please try again.`,
        timestamp: new Date(),
      });
    }

    return responses;
  }

  // Handle function calls recursively for compositional operations
  private async handleFunctionCalls(functionCalls: unknown[], depth: number = 0): Promise<ChatMessage[]> {
    const responses: ChatMessage[] = [];
    
    // Prevent infinite recursion - max 10 levels for complex styling operations
    // if (depth >= 100) {
    //   console.warn('⚠️ Maximum function call depth reached, stopping recursion');
    //   return responses;
    // }
    
    // Log function calls for debugging complex operations
    console.log(`📋 Executing ${functionCalls.length} functions at depth ${depth}:`, 
      (functionCalls as Array<{name: string}>).map(fc => fc.name).join(', '));
    
    // Execute all functions in parallel
    const functionPromises = (functionCalls as Array<{name: string; args: Record<string, unknown>}>).map(async (functionCall) => {
      const handler = this.functionHandlers.get(functionCall.name);
      
      if (handler) {
        try {
          const functionResult = await handler(functionCall.args as Record<string, unknown>);
          console.log(`✅ ${functionCall.name} completed successfully`);
          return {
            success: true,
            name: functionCall.name,
            result: functionResult,
          };
        } catch (error) {
          console.error(`❌ ${functionCall.name} failed:`, error);
          return {
            success: false,
            name: functionCall.name,
            error: String(error),
          };
        }
      } else {
        return {
          success: false,
          name: functionCall.name,
          error: `Unknown function: ${functionCall.name}`,
        };
      }
    });
    
    const results = await Promise.all(functionPromises);
    const functionResponseParts: unknown[] = [];
    const imagePartsToSend: unknown[] = [];
    
    results.forEach((result) => {
      if (result.success) {
        responses.push({
          id: this.generateId(),
          role: 'function',
          content: getFriendlyCompletionMessage(result.name),
          timestamp: new Date(),
          functionResponse: {
            name: result.name,
            result: result.result,
          },
        });
        
        functionResponseParts.push({
          functionResponse: {
            name: result.name,
            response: result.result,
          },
        });
        if (
          result.name === 'scan_sheet' &&
          result.result &&
          typeof result.result === 'object' &&
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (result.result as any).imageDataUrl &&
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          typeof (result.result as any).imageDataUrl === 'string'
        ) {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const dataUrl = (result.result as any).imageDataUrl as string;
          const commaIdx = dataUrl.indexOf(',');
          const base64Data = commaIdx >= 0 ? dataUrl.substring(commaIdx + 1) : dataUrl;
          imagePartsToSend.push({
            inlineData: {
              mimeType: 'image/png',
              data: base64Data,
            },
          });
        }
      } else {
        responses.push({
          id: this.generateId(),
          role: 'function',
          content: `Error executing ${result.name}: ${result.error}`,
          timestamp: new Date(),
          functionResponse: {
            name: result.name,
            result: null,
            error: result.error,
          },
        });
      }
    });
    
    // Send results back and check for more function calls
    if (functionResponseParts.length > 0) {
      try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        let followUpResult = await this.chat.sendMessage(functionResponseParts as any);
        let followUpResponse = followUpResult.response;
        if (imagePartsToSend.length > 0) {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          followUpResult = await this.chat.sendMessage(imagePartsToSend as any);
          followUpResponse = followUpResult.response;
        }
        
        const additionalCalls = followUpResponse.functionCalls();
        if (additionalCalls && additionalCalls.length > 0) {
          console.log(`🔄 Chaining ${additionalCalls.length} more functions (depth: ${depth + 1})`);
          const additionalMessages = await this.handleFunctionCalls(additionalCalls, depth + 1);
          responses.push(...additionalMessages);
        } else {
          const followUpText = followUpResponse.text();
          if (followUpText) {
            responses.push({
              id: this.generateId(),
              role: 'assistant',
              content: followUpText,
              timestamp: new Date(),
            });
          }
        }
      } catch (error) {
        console.error('Error in compositional function handling:', error);
      }
    }
    
    return responses;
  }

  // Format chat history for Gemini
  private formatHistory(messages: ChatMessage[]): Array<{role: string; parts: Array<{text: string}>}> {
    const filtered = messages
      .filter(msg => msg.role !== 'system' && msg.role !== 'function')
      .map(msg => ({
        role: msg.role === 'user' ? 'user' : 'model',
        parts: [{ text: msg.content }],
      }));
    
    // Gemini requires the first message to be from the user
    // If the first message is from the model, remove it
    if (filtered.length > 0 && filtered[0].role === 'model') {
      return filtered.slice(1);
    }
    
    return filtered;
  }

  // Generate unique ID for messages
  private generateId(): string {
    return `msg_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  }

  // Send with Gemini thinking via REST (config.thinkingConfig)
  private async sendWithThinking(
    formattedHistory: Array<{ role: string; parts: Array<{ text: string }> }> ,
    userParts: Array<Record<string, unknown>>
  ): Promise<ChatMessage[]> {
    const responses: ChatMessage[] = [];
    try {
      const body: Record<string, unknown> = {
        // Maintain roles for REST
        contents: [
          ...formattedHistory.map(h => ({ role: h.role, parts: h.parts })),
          { role: 'user', parts: userParts.length > 0 ? userParts : [{ text: '' }] },
        ],
        systemInstruction: { parts: [{ text: systemInstruction }] },
        tools: [ { functionDeclarations: allFunctionDeclarations } ],
        toolConfig: { functionCallingConfig: { mode: FunctionCallingMode.AUTO } },
        generationConfig: {
          temperature: 0.5,
          topP: 0.9,
          topK: 20,
          maxOutputTokens: 2048,
          candidateCount: 1,
          thinkingConfig: {
            thinkingBudget: -1,
            includeThoughts: true,
          },
        },
      };

      const url = `https://generativelanguage.googleapis.com/v1beta/models/${this.modelId}:generateContent`;
      const resp = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-goog-api-key': this.apiKey,
        },
        body: JSON.stringify(body),
      });
      if (!resp.ok) {
        const text = await resp.text();
        throw new Error(text);
      }
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const data: any = await resp.json();

      // If thoughts are included, we may receive thought parts; we keep them internal
      const assistantText = data?.candidates?.[0]?.content?.parts?.map((p: { text?: string }) => p.text).filter(Boolean).join('') || '';
      if (assistantText) {
        responses.push({
          id: this.generateId(),
          role: 'assistant',
          content: assistantText,
          timestamp: new Date(),
        });
      }
      return responses;
    } catch (error) {
      responses.push({
        id: this.generateId(),
        role: 'assistant',
        content: `I encountered an error: ${error}`,
        timestamp: new Date(),
      });
      return responses;
    }
  }

  // Best-effort: append grounding links from response if available
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private appendGroundingCitationsSafely(resp: any, baseText: string): string {
    try {
      // Gemini responses often expose grounding metadata on the top candidate
      const candidate = resp?.candidates?.[0];
      const gm = candidate?.groundingMetadata || candidate?.safetyRatings?.groundingMetadata || resp?.groundingMetadata;
      // Try common shapes
      const urls: string[] = [];
      const addUrl = (u: unknown) => {
        if (typeof u === 'string' && u.startsWith('http')) urls.push(u);
      };
      if (gm?.webSearchQueries && Array.isArray(gm.webSearchQueries)) {
        // ignore queries
      }
      if (gm?.searchResults && Array.isArray(gm.searchResults)) {
        gm.searchResults.forEach((r: unknown) => {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const rr = r as any;
          addUrl(rr?.url || rr?.link || rr?.uri);
        });
      }
      if (gm?.supportingUrls && Array.isArray(gm.supportingUrls)) {
        gm.supportingUrls.forEach(addUrl);
      }
      if (gm?.citations && Array.isArray(gm.citations)) {
        gm.citations.forEach((c: unknown) => {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const cc = c as any;
          addUrl(cc?.uri || cc?.url || cc?.link);
        });
      }
      const unique = Array.from(new Set(urls)).slice(0, 3);
      if (unique.length === 0) return baseText;
      const sources = unique.map((u, i) => `${i + 1}. ${u}`).join('\n');
      return `${baseText}\n\nSources:\n${sources}`;
    } catch {
      return baseText;
    }
  }

  // Reset the chat session
  public resetChat() {
    this.chat = this.model.startChat({
      history: [],
    });
  }

  // Update API key
  public updateApiKey(apiKey: string) {
    this.genAI = new GoogleGenerativeAI(apiKey);
    this.initializeModel();
  }
}
