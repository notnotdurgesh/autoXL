import { useState, useCallback } from 'react';

// Command interface for undo/redo operations
export interface Command {
  execute: () => void;
  undo: () => void;
  description: string;
}

export interface UndoRedoState {
  canUndo: boolean;
  canRedo: boolean;
  undoDescription: string;
  redoDescription: string;
}

export const useUndoRedo = () => {
  const [undoStack, setUndoStack] = useState<Command[]>([]);
  const [redoStack, setRedoStack] = useState<Command[]>([]);

  // Execute a command and add it to the undo stack
  const executeCommand = useCallback((command: Command) => {
    // Execute the command
    command.execute();
    
    // Add to undo stack
    setUndoStack(prev => [...prev, command]);
    
    // Clear redo stack when a new command is executed
    setRedoStack([]);
  }, []);

  // Undo the last command
  const undo = useCallback(() => {
    if (undoStack.length === 0) return;

    const command = undoStack[undoStack.length - 1];
    
    // Execute the undo
    command.undo();
    
    // Move command from undo stack to redo stack
    setUndoStack(prev => prev.slice(0, -1));
    setRedoStack(prev => [...prev, command]);
  }, [undoStack]);

  // Redo the last undone command
  const redo = useCallback(() => {
    if (redoStack.length === 0) return;

    const command = redoStack[redoStack.length - 1];
    
    // Execute the command again
    command.execute();
    
    // Move command from redo stack to undo stack
    setRedoStack(prev => prev.slice(0, -1));
    setUndoStack(prev => [...prev, command]);
  }, [redoStack]);

  // Clear all history
  const clearHistory = useCallback(() => {
    setUndoStack([]);
    setRedoStack([]);
  }, []);

  // Get current state
  const getState = useCallback((): UndoRedoState => {
    return {
      canUndo: undoStack.length > 0,
      canRedo: redoStack.length > 0,
      undoDescription: undoStack.length > 0 ? undoStack[undoStack.length - 1].description : '',
      redoDescription: redoStack.length > 0 ? redoStack[redoStack.length - 1].description : ''
    };
  }, [undoStack, redoStack]);

  return {
    executeCommand,
    undo,
    redo,
    clearHistory,
    getState,
    ...getState()
  };
};