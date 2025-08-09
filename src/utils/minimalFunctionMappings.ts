// Minimal friendly names for AI functions
export const minimalFunctionNames: Record<string, string> = {
  // Data Operations
  scan_sheet: 'Reading sheet',
  get_cell_value: 'Reading cell',
  get_cell_data: 'Reading cell data',
  set_cell_value: 'Editing cell',
  get_range_values: 'Reading range',
  set_range_values: 'Editing range',
  
  // Formatting
  set_cell_formatting: 'Applying styles',
  set_cell_formula: 'Adding formula',
  auto_fill: 'Auto-filling',
  
  // Structure
  insert_rows: 'Adding rows',
  insert_columns: 'Adding columns',
  delete_rows: 'Removing rows',
  delete_columns: 'Removing columns',
  
  // Data Operations
  sort_range: 'Sorting data',
  filter_range: 'Filtering data',
  find_replace: 'Finding & replacing',
  create_chart: 'Creating chart',
  
  // Cell Operations
  find_cell: 'Locating cell',
  clear_cell: 'Clearing cell',
  delete_cell: 'Deleting cell',
  clear_range: 'Clearing range',
  merge_range: 'Merging cells',
  unmerge_range: 'Unmerging cells',
  add_hyperlink: 'Adding link',
  
  // Analysis
  calculate_sum: 'Calculating sum',
  calculate_average: 'Calculating average',
  
  // Creative Operations
  copy_cells: 'Copying cells',
  move_cells: 'Moving cells',
  reorganize_data: 'Reorganizing',
  swap_cells: 'Swapping cells',
  transform_data: 'Transforming data',
  create_pattern: 'Creating pattern',
};

// Friendly completion messages for functions
export const friendlyCompletionMessages: Record<string, string> = {
  // Data Operations
  scan_sheet: 'Read the sheet',
  get_cell_value: 'Got the cell value',
  get_cell_data: 'Read cell details',
  set_cell_value: 'Edited the cell',
  get_range_values: 'Read the data',
  set_range_values: 'Updated the cells',
  
  // Formatting
  set_cell_formatting: 'Styled the cells',
  set_cell_formula: 'Added formula',
  auto_fill: 'Filled cells',
  
  // Structure
  insert_rows: 'Added rows',
  insert_columns: 'Added columns',
  delete_rows: 'Deleted rows',
  delete_columns: 'Deleted columns',
  
  // Data Operations
  sort_range: 'Sorted data',
  filter_range: 'Filtered data',
  find_replace: 'Replaced text',
  create_chart: 'Made a chart',
  
  // Cell Operations
  find_cell: 'Found the cell',
  clear_cell: 'Cleared cell',
  delete_cell: 'Deleted cell',
  clear_range: 'Cleared cells',
  merge_range: 'Merged cells',
  unmerge_range: 'Split cells',
  add_hyperlink: 'Added link',
  
  // Analysis
  calculate_sum: 'Calculated total',
  calculate_average: 'Got average',
  
  // Creative Operations
  copy_cells: 'Copied cells',
  move_cells: 'Moved cells',
  reorganize_data: 'Reorganized sheet',
  swap_cells: 'Swapped cells',
  transform_data: 'Transformed data',
  create_pattern: 'Made pattern',
};

export const getMinimalFunctionName = (functionName: string): string => {
  return minimalFunctionNames[functionName] || functionName.replace(/_/g, ' ');
};

export const getFriendlyCompletionMessage = (functionName: string): string => {
  return friendlyCompletionMessages[functionName] || `Done`;
};
