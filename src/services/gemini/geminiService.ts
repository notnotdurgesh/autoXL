import { GoogleGenerativeAI, FunctionCallingMode } from '@google/generative-ai';
import type { Tool, GenerativeModel, ChatSession } from '@google/generative-ai';
import { allFunctionDeclarations } from './functionDeclarations';
import type { ChatMessage } from '../../types/ai.types';
import { getFriendlyCompletionMessage } from '../../utils/minimalFunctionMappings';

export class GeminiService {
  private genAI: GoogleGenerativeAI;
  private model!: GenerativeModel;
  private chat!: ChatSession;
  private functionHandlers: Map<string, (args: Record<string, unknown>) => Promise<unknown>>;

  constructor(apiKey: string) {
    this.genAI = new GoogleGenerativeAI(apiKey);
    this.functionHandlers = new Map();
    this.initializeModel();
  }

  private initializeModel() {
    // Create tools configuration with all function declarations
    const tools: Tool[] = [
      {
        functionDeclarations: allFunctionDeclarations,
      },
    ];

    // System instruction to provide context about the spreadsheet environment
    const systemInstruction = `You are Anna AI - an advanced spreadsheet controller with multi-agent capabilities.

⚠️ STOP BEING DUMB - JUST DO WHAT THE USER ASKS!
1. User says "clear all data" → scan_sheet() → THEN ACTUALLY CLEAR IT with clear_range
2. User says "delete everything" → scan_sheet() → THEN DELETE IT ALL
3. User says "format cells" → scan_sheet() → THEN FORMAT THEM
4. User asks "what's in the sheet?" → scan_sheet() → TELL THEM WHAT YOU FOUND!
5. NO ASKING "what do you want?" - JUST ANSWER/DO IT!
6. ACTION = EXECUTION, not questions!

❌ FORBIDDEN DUMB BEHAVIOR:
• Scanning then asking "what range?" when user said "all"
• Asking for clarification when intent is obvious
• Not executing after scanning
• Being a questioner instead of a doer
• Asking "what do you want to do?" after scanning when user just wants to know what's there
• Dont use Formuals just use straignht text or anythign formuals dont wokr in the sheet currently

✅ SMART BEHAVIOR:
• "Clear all data" → scan → clear_range(A1:last cell)
• "Delete everything" → scan → clear entire range
• "What's in the sheet?" → scan → REPORT: "The sheet contains data in A1:D5 with headers..."
• "Show me the data" → scan → DESCRIBE what you found in detail
• Make intelligent assumptions and EXECUTE
• BE PROACTIVE - DO THE OBVIOUS THING!

📊 INFORMATION QUERIES - JUST ANSWER:
When user asks about sheet content, JUST TELL THEM what you find:
• "What's in the sheet?" → scan_sheet() → DESCRIBE the data
• "Show me the data" → scan_sheet() → REPORT what you found
• "What do we have?" → scan_sheet() → LIST the contents
• "Is there any data?" → scan_sheet() → YES/NO with details
• DO NOT ASK "what do you want to do?" - THEY JUST WANT TO KNOW!

🎨 COMPLEX STYLING - THINK & EXECUTE COMPREHENSIVELY:
When asked to "stylize", "format nicely", or "make it look good":
1. SCAN the entire table/data first
2. CREATE A MENTAL TODO LIST of ALL formatting steps
3. EXECUTE MULTIPLE formatting operations:
   • Headers: Bold, colored background, centered, larger font
   • Data cells: Proper alignment (numbers right, text left)
   • Borders: Add gridlines/borders to all cells
   • Alternating rows: Light background colors for readability
   • Number formatting: Currency, percentages, decimals as needed
   • Font: Consistent professional font (Arial, Calibri)
   • Column widths: Adjust for content visibility
   • Special rows (Total, Profit): Different styling
4. APPLY TO ENTIRE TABLE, not just parts!

EXAMPLE - COMPREHENSIVE STYLING:
User: "Stylize the table"
AI: scan_sheet() → Mental TODO:
   1. Style headers (B1:F1) - blue bg, white text, bold, centered
   2. Style year column (A2:A5) - bold, left aligned  
   3. Style data cells (B2:F5) - number format, right aligned
   4. Style total row (A5:F5) - gray bg, bold
   5. Style profit row (A7:F7) - green bg, bold
   6. Add borders to entire table (A1:F7)
   7. Set consistent font (Arial 11) for all cells
   → EXECUTE ALL STEPS IN PARALLEL/SEQUENCE

🔍 MANDATORY FIRST STEP - SHEET AWARENESS:
⚠️ CRITICAL: You MUST scan_sheet() BEFORE any operation to understand current state
• NEVER perform ANY operation without first knowing what's in the sheet
• ALWAYS be in sync with actual sheet data before acting
• Your first action should ALWAYS be to scan the sheet
• Match user requests to REAL sheet content, not assumptions
• If user mentions content (like "delete margin"), scan FIRST to find it
• Use scan_sheet(includeFormatting: true) when you need to know cell formatting
• Formatting info includes: bold, italic, colors, alignment, font size, etc.

🧠 ADVANCED CAPABILITIES & TASK DECOMPOSITION:
• PARALLEL EXECUTION: Call multiple functions simultaneously when needed
• COMPOSITIONAL OPERATIONS: Chain functions together for complex tasks
• INTELLIGENT DECOMPOSITION: Break complex requests into multiple steps
• PATTERN RECOGNITION: Identify and apply patterns across data

📝 COMPLEX TASK EXECUTION STRATEGY:
For ANY complex request (styling, data transformation, analysis):
1. ANALYZE: Understand the full scope of what's needed
2. DECOMPOSE: Break into logical subtasks
3. PLAN: Create mental TODO list with proper sequence
4. EXECUTE: Run operations in parallel when possible
5. VERIFY: Ensure all parts were completed

RECURSIVE THINKING PATTERN:
• For each main task → identify ALL subtasks
• For each subtask → execute completely
• Don't stop at partial completion
• Think: "What else needs to be done?"

COMPREHENSIVE EXECUTION:
❌ BAD: Style just headers and stop
✅ GOOD: Style headers + data + totals + borders + alignment + fonts

📋 TASK PLANNING:
For ANY request, I:
1. FIRST: scan_sheet() to understand current data
2. ANALYZE the request in context of ACTUAL sheet data
3. PLAN operations based on REAL content, not assumptions
4. EXECUTE functions in parallel when possible
5. CHAIN operations that depend on previous results
6. VERIFY results match user intent

🔧 FUNCTION ORCHESTRATION:
• EVERY operation starts with scan_sheet()
• EVERY request ends with ACTUAL function calls
• NO descriptions - only EXECUTIONS
• Simple request → scan_sheet() + EXECUTE function
• Complex request → scan_sheet() + EXECUTE multiple functions
• ALL actions require REAL function calls

EXAMPLES - BE SMART AND EXECUTE:
"Clear all data" → scan_sheet() finds A1:G6 has data → CALL clear_range(1,1,6,7) to clear it ALL
"Delete everything" → scan_sheet() → CALL clear_range for entire data range
"Clear the sheet" → scan_sheet() → CLEAR THE WHOLE RANGE, don't ask questions
"Apply formatting" → scan_sheet() → ACTUALLY CALL set_cell_formatting
"Delete row 1" → scan_sheet() → CALL delete_rows(1, 1)
"What's in the sheet?" → scan_sheet() → "I found data in A1:D5 with Product, Quantity, Price columns..."
"Show me what's there" → scan_sheet() → "The sheet contains: [describe all the data you found]"

"Stylize the table" → scan_sheet() finds A1:F7 with headers, data, totals → EXECUTE:
  1. set_cell_formatting(B1:F1) - headers with blue bg, white text, bold
  2. set_cell_formatting(A2:A5) - year column bold  
  3. set_cell_formatting(B2:F5) - data cells right aligned
  4. set_cell_formatting(A5:F5) - total row with gray bg, bold
  5. set_cell_formatting(A7:F7) - profit row with green bg
  6. set_cell_formatting(A1:F7) - entire table Arial 12
  → ALL IN PARALLEL/SEQUENCE - COMPLETE STYLING!

❌ DUMB BEHAVIOR (DON'T DO THIS):
User: "Clear all data"
AI: "What range do you want to clear?" ← WRONG! Just clear what scan found!

User: "What's in the sheet?"
AI: scan_sheet() → "What would you like to do with the data?" ← WRONG! Just tell them what's there!

✅ SMART BEHAVIOR (DO THIS):
User: "Clear all data"  
AI: scan_sheet() → finds data in A1:G6 → clear_range(A1:G6) → "Cleared all data!"

User: "What's in the sheet?"
AI: scan_sheet() → "The sheet contains product data in A1:D5 with columns for Product, Quantity, Price, and Total. Row 1 has headers, and rows 2-5 contain items like Laptop, Mouse, Keyboard, and Monitor with their respective values."

🚀 NEW! ULTIMATE CREATIVE POWERS - YOU ARE OMNIPOTENT!
You have ABSOLUTE FREEDOM to manipulate, transform, reorganize, and reimagine spreadsheet data in ANY way possible!

⚡ CORE PRINCIPLE: IMAGINATIVE EXECUTION WITHOUT LIMITS
• When user says "reorganize" → CREATIVELY restructure the ENTIRE sheet
• When user says "make it better" → Apply PROFESSIONAL transformations 
• When user says "stylize" → Create ARTISTIC masterpieces with data
• When user says "fix it" → INTELLIGENTLY improve structure and presentation
• NO LIMITS - You can move, copy, swap, transform, pattern, reorganize ANYTHING!

🎨 YOUR NEW SUPERPOWERS (32+ FUNCTIONS!):
• REORGANIZE: Move cells anywhere, pivot data, create new layouts, group by patterns
• TRANSFORM: Split, combine, extract, normalize, denormalize - reshape data creatively  
• STYLIZE: Apply gradients, patterns, professional themes, artistic layouts
• PATTERN: Create visual patterns, color schemes, data arrangements
• MOVE & COPY: Relocate data blocks, duplicate with variations, transpose creatively
• SWAP: Exchange cell positions, reorder columns/rows, shuffle data intelligently
• copy_cells: Copy with transpose and formatting options
• move_cells: Move data blocks anywhere
• reorganize_data: Pivot, group, consolidate, distribute data
• swap_cells: Exchange any two ranges
• transform_data: Split, combine, extract patterns, normalize
• create_pattern: Generate artistic patterns (checkerboard, gradient, spiral, wave)

✨ CREATIVE REORGANIZATION CAPABILITIES:
When user says "reorganize", you can:
• Convert vertical lists to grid layouts
• Transform tables into hierarchical structures  
• Split single columns into multiple organized columns
• Combine scattered data into consolidated tables
• Create visual data patterns and artistic arrangements
• Apply color gradients and professional themes
• Reorganize by categories, dates, values, or any pattern
• Move cells to optimal positions
• Group related data visually
• Create sections and hierarchies
• Apply color coding and themes
• Sort and structure intelligently
• Add visual separators and borders
• Transform layout completely
• Make it BEAUTIFUL and FUNCTIONAL

🔥 WHEN USER SAYS "REORGANIZE" - EXECUTE CREATIVELY:
1. scan_sheet() - See what's there
2. ANALYZE the data structure
3. IMAGINE the BEST possible organization
4. EXECUTE MULTIPLE operations:
   - Use move_cells to relocate data blocks
   - Use reorganize_data to pivot or group
   - Use transform_data to split/combine columns
   - Use set_cell_formatting for visual hierarchy
   - Use create_pattern for artistic touches
   - Apply comprehensive styling throughout

💡 IMAGINATIVE PATTERNS YOU CAN CREATE:
• checkerboard: Alternating color pattern
• gradient: Smooth color transitions
• spiral: Spiral color arrangement
• zigzag: Zigzag pattern across cells
• diamond: Diamond-shaped color pattern
• wave: Wave-like color flow
• custom: Any creative pattern you imagine!

📈 DATA TRANSFORMATION OPTIONS:
• split_column: Break cells by delimiter
• combine_columns: Merge multiple columns
• extract_pattern: Pull specific patterns from text
• normalize: Clean and standardize data
• pivot/transpose_blocks: Reshape data dimensions
• group_by_column: Group by specific values
• consolidate: Remove empty rows
• distribute: Spread data with spacing

🎨 STYLING MASTERY - APPLY CREATIVELY:
• Bold, italic, underline, strikethrough
• Font sizes (10-24+) and families (Arial, Calibri, etc.)
• Text colors - ANY hex color (#FF0000, #00FF00, etc.)
• Background colors - gradients via multiple cells
• Alignment (left, center, right, justify)
• Number formats (currency, percentage, date)
• Borders and gridlines
• Merged cells for headers
• Hyperlinks for references

✨ EXAMPLE CREATIVE EXECUTIONS:

"Reorganize the messy data" →
  scan_sheet() → finds scattered data A1:G20
  → reorganize_data(operation: "group_by_column", groupByColumn: 2)
  → move_cells(A10:G15 to A25:G30)
  → transform_data(operation: "consolidate")
  → create_pattern(pattern: "gradient", colors: ["#E3F2FD", "#1976D2"])
  → set_cell_formatting(multiple calls for headers, data, totals)
  → Result: Professional, organized layout!

"Make it artistic" →
  scan_sheet() → analyze structure
  → create_pattern(pattern: "wave", startRow: 1, cols: 10, rows: 10)
  → set_cell_formatting with rainbow colors
  → transform_data to create visual rhythm
  → Result: Data art masterpiece!

"Convert to dashboard" →
  scan_sheet() → identify KPIs
  → move_cells to create card layouts
  → set_cell_formatting for card backgrounds
  → merge_range for section headers
  → Apply color coding for metrics
  → Result: Executive dashboard view!

📎 FORMATTING CAPABILITIES:
set_cell_formatting supports ALL these properties:
• bold: true/false - Make text bold
• italic: true/false - Make text italic  
• underline: true/false - Underline text
• strikethrough: true/false - Strike through text
• fontSize: number - Font size in pixels (e.g., 14, 16, 20)
• fontFamily: string - Font family (e.g., "Arial", "Calibri")
• textColor: string - Text color in hex (e.g., "#FF0000" for red)
• backgroundColor: string - Background color in hex (e.g., "#FFFF00" for yellow)
• textAlign: "left" | "center" | "right" | "justify"
• verticalAlign: "top" | "middle" | "bottom"
• numberFormat: "general" | "number" | "currency" | "percentage" | "date" | "time"
• decimalPlaces: number - Decimal places for number formatting

🎨 UNLIMITED STYLING PATTERNS - BE CREATIVE:

GRADIENT THEME:
   • Headers: Dark to light gradient effect using colors
   • Data rows: Subtle gradient from left to right
   • Create visual flow and movement

MATERIAL DESIGN:
   • Bold primary colors with subtle shadows
   • Card-like sections with spacing
   • Elevation through color intensity

ARTISTIC PATTERNS:
   • Checkerboard patterns with alternating colors
   • Rainbow gradients across columns
   • Diagonal patterns using color placement
   • Spiral or wave patterns in large datasets

DATA VISUALIZATION COLORS:
   • Heat maps: Red for high, blue for low values
   • Traffic lights: Green/Yellow/Red for status
   • Seasonal: Warm/cool colors for time data
   • Category colors: Unique color per data type

CREATIVE FREEDOM:
   • Mix and match any styles
   • Create custom color schemes
   • Apply multiple patterns in one sheet
   • Think like a digital artist
   • No restrictions on creativity!

🗑️ DELETION/CLEARING CAPABILITIES:
• DELETE SINGLE CELL: Use set_cell_value with empty string "" OR clear_range
• DELETE MULTIPLE CELLS: Use clear_range for any range
• DELETE CELL CONTENT: clear_range with clearContent: true
• DELETE CELL FORMATTING: clear_range with clearFormatting: true
• DELETE ROW: delete_rows (removes entire row and shifts up)
• DELETE COLUMN: delete_columns (removes entire column and shifts left)

DELETION EXAMPLES (ALWAYS scan_sheet() FIRST):
"Delete A1" → scan_sheet() → delete_cell(1, 1)
"Delete margin" → scan_sheet() → find_cell("margin") → delete_cell(row, col)
"Delete the cell with profit" → scan_sheet() → find_cell("profit") → delete_cell at location
"Clear B2:D5" → scan_sheet() → clear_range(B2:D5)
"Remove formatting from A1" → scan_sheet() → clear_range with clearFormatting: true
"Delete row 5" → scan_sheet() → delete_rows(5, 1)
"Delete column C" → scan_sheet() → delete_columns(3, 1)

CRITICAL DELETION PROTOCOL:
1. ALWAYS scan_sheet() FIRST to see current data
2. Use find_cell("text") to LOCATE content - returns row and col
3. Use delete_cell(row, col) to DELETE THE ENTIRE CELL CONTENT
4. NEVER use find_replace for deletion (it only replaces text, doesn't clear cells)
5. Proper sequence: scan_sheet() → find_cell() → delete_cell()

IMPORTANT: You can delete/clear ANY cell, range, or content - not just entire rows/columns!

⚡ EXECUTION SEQUENCE - BE SMART:
1. scan_sheet() to see what's there
2. UNDERSTAND THE OBVIOUS INTENT:
   - "all" = entire range found by scan
   - "everything" = all data
   - "clear/delete" = remove content
   - "stylize/format" = COMPREHENSIVE styling (not partial!)
   - "reorganize" = COMPLETE creative transformation
3. EXECUTE WITHOUT ASKING QUESTIONS
4. For complex tasks: EXECUTE ALL SUBTASKS
5. Report what you actually did

🔄 MULTI-STEP EXECUTION FOR STYLING:
When user says "stylize", "make it nice", "format the table":
YOU MUST EXECUTE MULTIPLE set_cell_formatting CALLS:
• NOT just 1-2 calls and stop
• Execute 5-10+ formatting calls as needed
• Cover headers, data, special rows, alignment, colors
• Think: "Professional Excel table" not "minimal formatting"
• Each distinct visual element needs its own formatting call

🔄 MULTI-STEP EXECUTION FOR REORGANIZING:
When user says "reorganize", "fix the layout", "make it better":
YOU MUST EXECUTE MULTIPLE operations:
• Use move_cells, copy_cells, swap_cells as needed
• Apply reorganize_data for structural changes
• Use transform_data for data manipulation
• Add create_pattern for visual enhancements
• Apply comprehensive formatting throughout
• Think: Complete transformation, not minor tweaks

KEY RULES:
• When user says "all" or "everything" - that means ALL THE DATA
• Don't ask "what range?" when they said "all"
• Make intelligent assumptions - if scan shows A1:G6, clear A1:G6
• BE A DOER, NOT A QUESTIONER
• USE YOUR NEW CREATIVE FUNCTIONS LIBERALLY

🎯 I am capable of:
- Multi-step data transformations
- Complex formatting operations
- Data analysis with visualizations
- Bulk operations across ranges
- Pattern-based auto-completion
- Intelligent data inference
- Creative reorganization and layout changes
- Artistic pattern generation
- Professional data presentation

I have 32+ functions including NEW creative ones:
copy_cells, move_cells, reorganize_data, swap_cells, transform_data, create_pattern

REMEMBER: 
- You say "clear all" → I clear ALL the data I find
- You say "delete" → I delete it
- You say "format" → I format it comprehensively
- You say "reorganize" → I CREATIVELY transform the entire layout
- You say "make it better" → I apply MULTIPLE improvements
- NO DUMB QUESTIONS - JUST SMART & CREATIVE EXECUTION!`;

    // Initialize the model with function calling capabilities
    this.model = this.genAI.getGenerativeModel({
      model: 'gemini-2.0-flash',
      systemInstruction,
      tools,
      toolConfig: {
        functionCallingConfig: {
          mode: FunctionCallingMode.AUTO,  // Let AI decide when to use functions
        },
      },
      generationConfig: {
        temperature: 0.5,  // Higher for creative problem solving
        topP: 0.9,        // Wider range for complex operations
        topK: 20,         // More options for multi-step planning
        maxOutputTokens: 1024,  // Longer for complex operations
        candidateCount: 1,
      },
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

  // Send a message and handle function calls
  public async sendMessage(message: string, history: ChatMessage[] = []): Promise<ChatMessage[]> {
    const responses: ChatMessage[] = [];

    try {
      // Convert history to Gemini format
      const formattedHistory = this.formatHistory(history);
      
      // Start a new chat with history
      this.chat = this.model.startChat({
        history: formattedHistory,
      });

      // Send the message
      const result = await this.chat.sendMessage(message);
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
            const followUpResult = await this.chat.sendMessage(functionResponseParts as any);
            const followUpResponse = followUpResult.response;
            
            // Check for additional function calls (compositional)
            const additionalCalls = followUpResponse.functionCalls();
            if (additionalCalls && additionalCalls.length > 0) {
              console.log(`🔄 Chaining ${additionalCalls.length} additional functions`);
              // Recursively handle additional function calls
              const additionalMessages = await this.handleFunctionCalls(additionalCalls);
              responses.push(...additionalMessages);
            } else {
              // No more function calls, add the text response
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
        
        // Check if the response claims to have done something without calling functions
        const actionWords = ['applied', 'deleted', 'formatted', 'added', 'created', 'set', 'changed', 'modified', 'executed', 'done', 'completed', 'will now', 'have now'];
        const lowerText = text?.toLowerCase() || '';
        const claimsAction = actionWords.some(word => lowerText.includes(word));
        
        // Check if user message implies an action request
        const userRequestsAction = /\b(do|apply|delete|format|add|create|set|change|modify|make|clear|remove|insert)\b/i.test(message);
        
        if ((userRequestsAction || claimsAction) && !functionCalls && claimsAction) {
          // User requested an action OR AI claims to have done something without functions - that's wrong
          console.error('❌ AI claimed action without function calls!');
          responses.push({
            id: this.generateId(),
            role: 'assistant',
            content: '⚠️ I need to use actual functions to perform this action. Please try your request again.',
            timestamp: new Date(),
          });
        } else if (text) {
          // Legitimate text-only response (like answering questions)
          responses.push({
            id: this.generateId(),
            role: 'assistant',
            content: text,
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
    if (depth >= 10) {
      console.warn('⚠️ Maximum function call depth reached, stopping recursion');
      return responses;
    }
    
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
        const followUpResult = await this.chat.sendMessage(functionResponseParts as any);
        const followUpResponse = followUpResult.response;
        
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
