/**
 * Centralized Format Manager
 * Handles all formatting logic for table cells in a DRY, scalable way
 */

export interface FormatState {
  [key: string]: boolean | string;
}

export const INLINE_FORMATS = ['bold', 'italic', 'underline', 'strike'] as const;
export const STRUCTURAL_FORMATS = ['font', 'size', 'color', 'background', 'align'] as const;
export const TABLE_FORMATS = ['table', 'table-cell', 'table-cell-block', 'cell', 'row', 'td', 'th', 'tr'] as const;

export type InlineFormat = typeof INLINE_FORMATS[number];
export type StructuralFormat = typeof STRUCTURAL_FORMATS[number];
export type TableFormat = typeof TABLE_FORMATS[number];

/**
 * Format Manager Class
 * Manages global format toggles and cell-specific formatting
 */
export class FormatManager {
  private globalFormats: FormatState = {};
  private deselectedFormats: FormatState = {};

  /**
   * Check if a format is an inline format
   */
  isInlineFormat(format: string): format is InlineFormat {
    return (INLINE_FORMATS as readonly string[]).includes(format);
  }

  /**
   * Check if a format is a structural format
   */
  isStructuralFormat(format: string): format is StructuralFormat {
    return (STRUCTURAL_FORMATS as readonly string[]).includes(format);
  }

  /**
   * Check if a format is a table-related format
   */
  isTableFormat(format: string): format is TableFormat {
    return (TABLE_FORMATS as readonly string[]).includes(format);
  }

  /**
   * Toggle a global format on/off
   * When enabled, format applies to all new/focused cells
   * When disabled, format is removed from global state
   * 
   * IMPORTANT: Font and size are value-based formats (not toggles)
   * - They are never added to deselectedFormats
   * - Changing their value replaces the old value
   * - They always have a value (no "deselected" state)
   */
  toggleGlobalFormat(format: string, value: boolean | string): void {
    // Check if this is a value-based format (font, size)
    const isValueBasedFormat = format === 'font' || format === 'size';
    
    if (value === false || value === null || value === undefined) {
      // Deactivating format
      delete this.globalFormats[format];
      
      // Only add to deselectedFormats if it's NOT a value-based format
      // Value-based formats (font, size) are never "deselected"
      if (!isValueBasedFormat) {
        this.deselectedFormats[format] = true;
        console.log(`🎨 FormatManager: Deactivated global format "${format}"`);
      } else {
        console.log(`🎨 FormatManager: Removed value-based format "${format}" (not marked as deselected)`);
      }
    } else {
      // Activating format
      this.globalFormats[format] = value;
      
      // Remove from deselectedFormats (format is now active)
      delete this.deselectedFormats[format];
      console.log(`🎨 FormatManager: Activated global format "${format}" with value:`, value);
    }
  }

  /**
   * Get all currently active global formats
   */
  getGlobalFormats(): FormatState {
    return { ...this.globalFormats };
  }

  /**
   * Get all deselected formats
   */
  getDeselectedFormats(): FormatState {
    return { ...this.deselectedFormats };
  }

  /**
   * Check if a format is globally active
   */
  isGloballyActive(format: string): boolean {
    return format in this.globalFormats && this.globalFormats[format] !== false;
  }

  /**
   * Check if a format is deselected
   */
  isDeselected(format: string): boolean {
    return this.deselectedFormats[format] === true;
  }

  /**
   * Merge global formats with cell-specific formats
   * Cell formats take precedence over global formats
   */
  mergeCellFormats(cellFormats: FormatState): FormatState {
    const merged: FormatState = {};

    // Start with global formats
    Object.keys(this.globalFormats).forEach(format => {
      if (this.isInlineFormat(format)) {
        merged[format] = this.globalFormats[format];
      }
    });

    // Override with cell-specific formats
    Object.keys(cellFormats).forEach(format => {
      if (!this.isTableFormat(format)) {
        merged[format] = cellFormats[format];
      }
    });

    return merged;
  }

  /**
   * Get formats to display in toolbar for a given cell
   * Shows cell's actual formats, not global formats
   */
  getToolbarFormats(cellFormats: FormatState): FormatState {
    const toolbarFormats: FormatState = {};

    // Show actual cell formats (cell state takes precedence)
    Object.keys(cellFormats).forEach(format => {
      if (!this.isTableFormat(format)) {
        toolbarFormats[format] = cellFormats[format];
      }
    });

    return toolbarFormats;
  }

  /**
   * Get formats to apply to a newly focused or created cell
   * Applies global formats only
   */
  getFormatsForNewCell(): FormatState {
    return this.getGlobalFormats();
  }

  /**
   * Filter out table-related formats from a format object
   */
  filterTableFormats(formats: FormatState): FormatState {
    const filtered: FormatState = {};
    Object.keys(formats).forEach(format => {
      if (!this.isTableFormat(format)) {
        filtered[format] = formats[format];
      }
    });
    return filtered;
  }

  /**
   * Filter out deselected formats from a format object
   * Only filters inline formats (structural formats like font/size persist)
   */
  filterDeselectedFormats(formats: FormatState): FormatState {
    const filtered: FormatState = { ...formats };
    
    Object.keys(this.deselectedFormats).forEach(format => {
      if (this.deselectedFormats[format] && this.isInlineFormat(format)) {
        delete filtered[format];
      }
    });

    return filtered;
  }

  /**
   * Clear all global formats
   */
  clearGlobalFormats(): void {
    this.globalFormats = {};
    console.log('🎨 FormatManager: Cleared all global formats');
  }

  /**
   * Clear all deselected formats
   */
  clearDeselectedFormats(): void {
    this.deselectedFormats = {};
    console.log('🎨 FormatManager: Cleared all deselected formats');
  }

  /**
   * Reset the format manager to initial state
   */
  reset(): void {
    this.clearGlobalFormats();
    this.clearDeselectedFormats();
    console.log('🎨 FormatManager: Reset to initial state');
  }
}
