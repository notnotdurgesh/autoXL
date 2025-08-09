// Friendly names and descriptions for all AI functions
export interface FunctionMapping {
  friendlyName: string;
  description: string;
  icon: string;
  category: 'data' | 'format' | 'structure' | 'analysis' | 'creative';
  color: string;
}

export const functionMappings: Record<string, FunctionMapping> = {
  // Data Operations
  scan_sheet: {
    friendlyName: '🔍 Scanning Sheet',
    description: 'Analyzing the spreadsheet content',
    icon: '🔍',
    category: 'analysis',
    color: '#3B82F6',
  },
  get_cell_value: {
    friendlyName: '📋 Reading Cell',
    description: 'Getting cell value',
    icon: '📋',
    category: 'data',
    color: '#10B981',
  },
  get_cell_data: {
    friendlyName: '📊 Reading Cell Details',
    description: 'Getting cell value and formatting',
    icon: '📊',
    category: 'data',
    color: '#10B981',
  },
  set_cell_value: {
    friendlyName: '✏️ Writing to Cell',
    description: 'Setting cell value',
    icon: '✏️',
    category: 'data',
    color: '#8B5CF6',
  },
  get_range_values: {
    friendlyName: '📑 Reading Range',
    description: 'Getting multiple cell values',
    icon: '📑',
    category: 'data',
    color: '#10B981',
  },
  set_range_values: {
    friendlyName: '📝 Writing Range',
    description: 'Setting multiple cell values',
    icon: '📝',
    category: 'data',
    color: '#8B5CF6',
  },
  
  // Formatting Operations
  set_cell_formatting: {
    friendlyName: '🎨 Styling Cells',
    description: 'Applying formatting and colors',
    icon: '🎨',
    category: 'format',
    color: '#EC4899',
  },
  set_cell_formula: {
    friendlyName: '🧮 Adding Formula',
    description: 'Setting cell formula',
    icon: '🧮',
    category: 'data',
    color: '#F59E0B',
  },
  auto_fill: {
    friendlyName: '🔄 Auto-Filling',
    description: 'Filling cells with pattern',
    icon: '🔄',
    category: 'data',
    color: '#14B8A6',
  },
  
  // Structure Operations
  insert_rows: {
    friendlyName: '➕ Adding Rows',
    description: 'Inserting new rows',
    icon: '➕',
    category: 'structure',
    color: '#06B6D4',
  },
  insert_columns: {
    friendlyName: '➕ Adding Columns',
    description: 'Inserting new columns',
    icon: '➕',
    category: 'structure',
    color: '#06B6D4',
  },
  delete_rows: {
    friendlyName: '➖ Removing Rows',
    description: 'Deleting rows',
    icon: '➖',
    category: 'structure',
    color: '#EF4444',
  },
  delete_columns: {
    friendlyName: '➖ Removing Columns',
    description: 'Deleting columns',
    icon: '➖',
    category: 'structure',
    color: '#EF4444',
  },
  
  // Data Manipulation
  sort_range: {
    friendlyName: '↕️ Sorting Data',
    description: 'Organizing data by values',
    icon: '↕️',
    category: 'data',
    color: '#6366F1',
  },
  filter_range: {
    friendlyName: '🔽 Filtering Data',
    description: 'Showing specific data only',
    icon: '🔽',
    category: 'data',
    color: '#6366F1',
  },
  find_replace: {
    friendlyName: '🔍 Find & Replace',
    description: 'Searching and replacing values',
    icon: '🔍',
    category: 'data',
    color: '#F97316',
  },
  
  // Visual Operations
  create_chart: {
    friendlyName: '📈 Creating Chart',
    description: 'Generating data visualization',
    icon: '📈',
    category: 'analysis',
    color: '#0EA5E9',
  },
  
  // Utility Operations
  find_cell: {
    friendlyName: '🎯 Locating Cell',
    description: 'Finding specific content',
    icon: '🎯',
    category: 'data',
    color: '#3B82F6',
  },
  clear_cell: {
    friendlyName: '🧹 Clearing Cell',
    description: 'Removing cell content',
    icon: '🧹',
    category: 'data',
    color: '#F87171',
  },
  delete_cell: {
    friendlyName: '🗑️ Deleting Cell',
    description: 'Removing cell content',
    icon: '🗑️',
    category: 'data',
    color: '#F87171',
  },
  clear_range: {
    friendlyName: '🧹 Clearing Range',
    description: 'Removing content from multiple cells',
    icon: '🧹',
    category: 'data',
    color: '#F87171',
  },
  merge_range: {
    friendlyName: '🔗 Merging Cells',
    description: 'Combining cells together',
    icon: '🔗',
    category: 'structure',
    color: '#A855F7',
  },
  unmerge_range: {
    friendlyName: '✂️ Unmerging Cells',
    description: 'Separating merged cells',
    icon: '✂️',
    category: 'structure',
    color: '#A855F7',
  },
  add_hyperlink: {
    friendlyName: '🔗 Adding Link',
    description: 'Creating hyperlink in cell',
    icon: '🔗',
    category: 'data',
    color: '#3B82F6',
  },
  
  // Analysis Operations
  calculate_sum: {
    friendlyName: '➕ Calculating Sum',
    description: 'Adding up values',
    icon: '➕',
    category: 'analysis',
    color: '#10B981',
  },
  calculate_average: {
    friendlyName: '📊 Calculating Average',
    description: 'Finding average value',
    icon: '📊',
    category: 'analysis',
    color: '#10B981',
  },
  
  // NEW Creative Operations
  copy_cells: {
    friendlyName: '📋 Copying Cells',
    description: 'Duplicating cell data creatively',
    icon: '📋',
    category: 'creative',
    color: '#FF6B6B',
  },
  move_cells: {
    friendlyName: '📦 Moving Cells',
    description: 'Relocating data blocks',
    icon: '📦',
    category: 'creative',
    color: '#4ECDC4',
  },
  reorganize_data: {
    friendlyName: '🎨 Reorganizing Layout',
    description: 'Transforming data structure creatively',
    icon: '🎨',
    category: 'creative',
    color: '#95E1D3',
  },
  swap_cells: {
    friendlyName: '🔄 Swapping Cells',
    description: 'Exchanging cell positions',
    icon: '🔄',
    category: 'creative',
    color: '#FFA07A',
  },
  transform_data: {
    friendlyName: '✨ Transforming Data',
    description: 'Applying creative data transformations',
    icon: '✨',
    category: 'creative',
    color: '#DDA0DD',
  },
  create_pattern: {
    friendlyName: '🎯 Creating Pattern',
    description: 'Generating artistic patterns',
    icon: '🎯',
    category: 'creative',
    color: '#FFD700',
  },
};

// Helper function to get friendly info for a function
export const getFunctionInfo = (functionName: string): FunctionMapping => {
  return functionMappings[functionName] || {
    friendlyName: `📌 ${functionName.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase())}`,
    description: 'Processing request',
    icon: '📌',
    category: 'data',
    color: '#9CA3AF',
  };
};

// Get category color
export const getCategoryColor = (category: FunctionMapping['category']): string => {
  const colors = {
    data: '#10B981',
    format: '#EC4899',
    structure: '#06B6D4',
    analysis: '#3B82F6',
    creative: '#FF6B6B',
  };
  return colors[category] || '#9CA3AF';
};

// Animation keyframes for different categories
export const getCategoryAnimation = (category: FunctionMapping['category']): string => {
  const animations = {
    data: 'pulse',
    format: 'bounce',
    structure: 'slide',
    analysis: 'spin',
    creative: 'glow',
  };
  return animations[category] || 'fade';
};
