import Quill from 'quill';
import Delta from 'quill-delta';
import type { EmitterSource, Range } from 'quill';
import type { Props } from './types';
import type { BindingObject, Context } from './types/keyboard';
import { FormatManager } from './utils/format-manager';
import {
  cellId,
  TableCellBlock,
  TableThBlock,
  TableCell,
  TableTh,
  TableRow,
  TableThRow,
  TableBody,
  TableThead,
  TableTemporary,
  TableContainer,
  tableId,
  TableCol,
  TableColgroup
} from './formats/table';
import TableHeader from './formats/header';
import { ListContainer } from './formats/list';
import { 
  matchTable,
  matchTableCell,
  matchTableCol,
  matchTableTemporary
} from './utils/clipboard-matchers';
import Language from './language';
import CellSelection from './ui/cell-selection';
import OperateLine from './ui/operate-line';
import TableMenus from './ui/table-menus';
import ToolbarTable, { TableSelect } from './ui/toolbar-table';
import { getCellId, getCorrectCellBlot } from './utils';
import TableToolbar from './modules/toolbar';
import TableClipboard from './modules/clipboard';

interface Options {
  language?: string | {
    name: string;
    content: Props;
  }
  menus?: string[]
  toolbarButtons?: {
    whiteList?: string[];
    singleWhiteList?: string[];
  }
  toolbarTable?: boolean;
}

type Line = TableCellBlock | TableHeader | ListContainer;

const Module = Quill.import('core/module');

class Table extends Module {
  language: Language;
  cellSelection: CellSelection;
  operateLine: OperateLine;
  tableMenus: TableMenus;
  tableSelect: TableSelect;
  options: Options;
  

  static keyboardBindings: { [propName: string]: BindingObject };
  formatManager: FormatManager;
  lastFormat: any = null;
  defaultFormats: any = {};
  lastClickedCell: HTMLElement | null = null;
  cellOriginTracker: WeakMap<any, { startedAsEmpty: boolean }> = new WeakMap();
  isPerformingTableOperation = false;
  isSettingUpTableCell = false;
  currentCell: any = null;
  
  static register() {
    Quill.register(TableCellBlock, true);
    Quill.register(TableCell, true);
    Quill.register(TableTh, true);
    Quill.register(TableRow, true);
    Quill.register(TableThRow, true);
    Quill.register(TableBody, true);
    Quill.register(TableThead, true);
    Quill.register(TableTemporary, true);
    Quill.register(TableContainer, true);
    Quill.register(TableCol, true);
    Quill.register(TableColgroup, true);
    Quill.register({
      [TableThBlock.blotName]: TableThBlock,
      'modules/toolbar': TableToolbar,
      'modules/clipboard': TableClipboard
    }, true);
  }
  constructor(quill: Quill, options: Options | any) {
    console.log('🔍🔍🔍 TABLE CONSTRUCTOR STARTED! 🔍🔍🔍');
    super(quill, options);
    this.options = options;
    quill.clipboard.addMatcher('td, th', matchTableCell);
    quill.clipboard.addMatcher('tr', matchTable);
    quill.clipboard.addMatcher('col', matchTableCol);
    quill.clipboard.addMatcher('table', matchTableTemporary);
    this.language = new Language(options?.language);
    this.cellSelection = new CellSelection(quill, this);
    this.operateLine = new OperateLine(quill, this);
    this.tableMenus = new TableMenus(quill, this);
    this.tableSelect = new TableSelect();
    console.log('🔍 CONSTRUCTOR: TableSelect created');
    this.defaultFormats = options.defaultFormats || {};
    this.formatManager = new FormatManager();
    console.log('🔍 CONSTRUCTOR: FormatManager initialized');
    
    // Initialize global formats with default values if provided
    if (this.defaultFormats && Object.keys(this.defaultFormats).length > 0) {
      Object.keys(this.defaultFormats).forEach(format => {
        if (this.formatManager.isStructuralFormat(format)) {
          this.formatManager.toggleGlobalFormat(format, this.defaultFormats[format]);
        }
      });
      console.log('🔍 INIT: Initialized global formats with defaults');
    }
    
    console.log('🔍 CONSTRUCTOR: About to add event listeners');
        
    // NOTE: Toolbar update filtering is now handled in updateToolbarUI() method
    // which respects the isExistingContent flag to only filter deselected formats
    // for new/empty cells, never for existing content
    console.log('🔍 CONSTRUCTOR: Setting up text-change listener');
   
    quill.root.addEventListener('keyup', this.handleKeyup.bind(this));
    quill.root.addEventListener('mousedown', this.handleMousedown.bind(this));
    quill.root.addEventListener('scroll', this.handleScroll.bind(this));
    this.listenDeleteTable();
    this.registerToolbarTable(options?.toolbarTable);
    
    let isInTable = false;

    // Track when user is in a table cell
    this.quill.on('selection-change', (range: any, oldRange: any, source: string) => {
      if (source === Quill.sources.SILENT) {
        return;
      }

      if (!range) {
        isInTable = false;
        this.currentCell = null;
        return;
      }

      const tableModule = this.quill.getModule('table-better');
      if (!tableModule) return;

      const [__, _, cell] = tableModule.getTable(range);

      if (!cell || !(cell instanceof HTMLElement)) {
        isInTable = false;
        this.currentCell = null;
        return;
      }

      isInTable = true;
      this.currentCell = cell;
    });

      // Handle text-change events - apply and persist formats
    this.quill.on('text-change', (delta: any, oldContents: any, source: string) => {
      // CRITICAL: Detect structural table changes (row/column insertion)
      // When a row or column is inserted, the delta will contain table structure operations
      // In these cases, we must skip ALL processing to avoid interfering with post-insert focus
      const isStructuralTableChange = delta.ops && delta.ops.some((op: any) => {
        if (op.insert && typeof op.insert === 'object') {
          return op.attributes && (op.attributes['table-cell'] || op.attributes['table-row'] || op.attributes['table']);
        }
        return false;
      });

      if (isStructuralTableChange) {
        console.log('🔍 TEXT-CHANGE (quill-table-better): Detected structural table change - skipping processing');
        return;
      }

      if (!isInTable || !this.currentCell) return;

      // Keep menu visible during all editor changes
      if (this.tableMenus && this.cellSelection.selectedTds.length > 0) {
        this.tableMenus.showMenus();
      }

      const range = this.quill.getSelection();
      if (!range) return;

      if (source === Quill.sources.USER) {
        // CRITICAL: Get current global formats and deselected formats
        const globalFormats = this.formatManager.getGlobalFormats();
        const deselectedFormats = this.formatManager.getDeselectedFormats();
        
        let position = 0;
        for (const op of delta.ops) {
          if (op.retain) {
            position += op.retain;
          } else if (op.insert && typeof op.insert === 'string') {
            // Check if current cell has existing content AFTER the new text is inserted
            const [__, ___, cell] = this.getTable(range);
            let hasExistingContent = false;
            let cellFormats: any = {};
            let startedAsEmpty = false;
            
            if (cell && cell.domNode) {
              const cellText = (cell.domNode as HTMLElement).innerText || '';
              const currentText = cellText.replace(/\u200B/g, '').trim();
              
              // PRIORITY 2 FIX: Check if cell was empty BEFORE this keystroke
              // A cell that had zero characters before is "empty" regardless of what's in it now
              const cellWasEmpty = currentText.length === 0;
              
              // Only treat as "existing content" if cell had content BEFORE this keystroke
              hasExistingContent = currentText.length > 0;
              
              const cellTracker = this.cellOriginTracker.get(cell);
              if (cellTracker !== undefined) {
                startedAsEmpty = cellTracker.startedAsEmpty;
              } else {
                startedAsEmpty = cellWasEmpty;
                this.cellOriginTracker.set(cell, { startedAsEmpty });
              }
              
              if (hasExistingContent) {
                if (startedAsEmpty) {
                // CRITICAL: Priority order for empty cell that's now getting content:
                // 1. FormatManager.globalFormats (HIGHEST - what user selected in toolbar)
                // 2. Previous cell formats (FALLBACK - for formats not in globalFormats)
                const globalFormats = this.formatManager.getGlobalFormats();
                const previousCellFormats = this.getPreviousCellFormats(this.currentCell);
                
                // Start with global formats (highest priority)
                cellFormats = { ...globalFormats };
                
                // Merge previous cell formats ONLY for formats NOT in globalFormats
                if (Object.keys(previousCellFormats).length > 0) {
                  Object.keys(previousCellFormats).forEach(format => {
                    if (previousCellFormats[format] !== undefined && !globalFormats[format]) {
                      cellFormats[format] = previousCellFormats[format];
                    }
                  });
                }
                } else {
                  const insertionIndex = position;
                  const formatsAtPosition = this.quill.getFormat(insertionIndex, 1);
                  cellFormats = this.formatManager.filterDeselectedFormats(formatsAtPosition);
                }
              }
            }
            
            // CRITICAL: Remove deselected formats from the newly inserted text
            const deselectedFormats = this.formatManager.getDeselectedFormats();
            const formatsToRemove: any = {};
            
            Object.keys(deselectedFormats).forEach(format => {
              if (deselectedFormats[format] && this.formatManager.isInlineFormat(format)) {
                formatsToRemove[format] = false;
              }
            });
            
            if (Object.keys(formatsToRemove).length > 0) {
              this.quill.formatText(position, op.insert.length, formatsToRemove, Quill.sources.SILENT);
              
              const [__, ___, cell] : any = this.getTable(range);
              if (cell) {
                const cellBlot = Quill.find(cell) as any;
                if (cellBlot) {
                  const cellIndex = this.quill.getIndex(cellBlot);
                  const cellLength = cellBlot.length() - 1;
                  
                  this.quill.formatText(cellIndex, cellLength, formatsToRemove, Quill.sources.SILENT);
                }
              }
            }
            
            // Apply formats based on cell content and origin
            let formatsToApply: any = {};
            let shouldFilterDeselected = true;
            
            if (hasExistingContent) {
              if (startedAsEmpty) {
                formatsToApply = this.formatManager.getGlobalFormats();
                formatsToApply = this.formatManager.filterDeselectedFormats(formatsToApply);
                shouldFilterDeselected = true;
              } else {
                formatsToApply = this.formatManager.filterDeselectedFormats(cellFormats);
                shouldFilterDeselected = false;
              }
            } else {
              // Empty cell: use existing formats or global formats
              const cellBlot = Quill.find(this.cellSelection.selectedTds[0]) as any;
              if (cellBlot) {
                const cellIndex = this.quill.getIndex(cellBlot);
                formatsToApply = this.quill.getFormat(cellIndex, 1);
              }
              
              // If no existing formats, use global formats
              if (Object.keys(formatsToApply).length === 0) {
                formatsToApply = this.formatManager.getGlobalFormats();
              }
              
              // Remove deselected formats
              formatsToApply = this.formatManager.filterDeselectedFormats(formatsToApply);
              shouldFilterDeselected = true;
            }
            
            // CRITICAL: Only apply formats if cell started empty
            // For cells with existing content, DO NOT override their format state
            if (startedAsEmpty && Object.keys(formatsToApply).length > 0) {
              formatsToApply = this.formatManager.filterDeselectedFormats(formatsToApply);
              this.quill.formatText(position, op.insert.length, formatsToApply, Quill.sources.SILENT);
              
              // IMPORTANT: Track the formats that were actually applied
              // This will be used for format inheritance in subsequent cells
              this.updateLastAppliedFormats(formatsToApply);
              
              if (Object.keys(formatsToRemove).length > 0) {
                setTimeout(() => {
                  this.quill.formatText(position, op.insert.length, formatsToRemove, Quill.sources.SILENT);
                  console.log('🔍 TEXT-CHANGE: Cleanup - removed deselected formats after applying');
                }, 5);
              }
            } else if (!startedAsEmpty) {
              console.log('🔍 TEXT-CHANGE: Cell had existing content - NOT applying any formats, preserving cell format state');
            }
            
            // Update toolbar UI
            const isExistingContentForToolbar = !shouldFilterDeselected;
            formatsToApply = this.formatManager.filterDeselectedFormats(formatsToApply);
            this.updateToolbarUI(formatsToApply);
            console.log('🔍 TEXT-CHANGE: Updated toolbar UI');
          }
          
          position += op.insert?.length || 0;
        }
      }
    });
  }

  clearHistorySelected() {
    const [table] = this.getTable();
    if (!table) return;
    const selectedTds: Element[] = Array.from(
      table.domNode.querySelectorAll('td.ql-cell-focused, td.ql-cell-selected')
    );
    for (const td of selectedTds) {
      td.classList && td.classList.remove('ql-cell-focused', 'ql-cell-selected');
    }
  }

  deleteTable() {
    const [table] = this.getTable();
    if (table == null) return;
    const offset = table.offset();
    table.remove();
    this.hideTools();
    this.quill.update(Quill.sources.USER);
    this.quill.setSelection(offset, Quill.sources.SILENT);
    // Check if editor is empty after deleting table
    const editorText = this.quill.getText().trim();
    if (editorText === '' || editorText === '\n') {
      // Editor is empty, set default font and size
      // Use stored defaults or fall back to reasonable value
    
        
      const defaultFont = this.defaultFormats?.font;
      const defaultSize = this.defaultFormats?.size;

      // Apply default formats at current cursor position
      // Use silent source for the first formatting to avoid multiple updates
      this.quill.format('font', defaultFont, Quill.sources.SILENT);
      this.quill.format('size', defaultSize, Quill.sources.USER); // USER source for the last one to trigger an update

      // Apply to a small range to ensure formats stick
      const range = this.quill.getSelection() || { index: 0, length: 1 };

      this.quill.formatText(range.index, Math.max(1, range.length), {
        font: defaultFont,
        size: defaultSize
      }, Quill.sources.USER);
      // Update format tracking variables to ensure further edits use these formats
      this.lastFormat = {
        ...this.lastFormat,
        font: defaultFont,
        size: defaultSize
      };
      console.log("this.lastFormat", this.lastFormat)
    }
  }

  deleteTableTemporary(source: EmitterSource = Quill.sources.API) {
    const temporaries = this.quill.scroll.descendants(TableTemporary);
    for (const temporary of temporaries) {
      temporary.remove();
    }
    this.hideTools();
    this.quill.update(source);
  }

  getTable(
    range = this.quill.getSelection()
  ): [null, null, null, -1] | [TableContainer, TableRow, TableCell, number] {
    if (range == null) return [null, null, null, -1];
    const [block, offset] = this.quill.getLine(range.index);
    if (block == null || block.statics.blotName !== TableCellBlock.blotName) {
      return [null, null, null, -1];
    }
    const cell = block.parent as TableCell;
    const row = cell.parent as TableRow;
    const table = row.parent.parent as TableContainer;
    return [table, row, cell, offset];
  }


  cellIsEmpty(cell: any): boolean {
    if (!cell) return false;
    const text = cell.domNode?.innerText || '';
    // Remove zero-width spaces and whitespace
    return text.replace(/\u200B/g, '').trim() === '';
  }

 updateToolbarUI(formats: any, isExistingContent: boolean = true) {
    const toolbarContainer = document.querySelector('.ql-toolbar');
    if (!toolbarContainer) return;
    
    // Create a copy to avoid modifying the original
    let filteredFormats = { ...formats };
    
    // CRITICAL: ONLY filter deselected formats for NEW/EMPTY cells
    // For existing content, ALWAYS show the actual formats of the cell
    // Deselection is a GLOBAL user preference that applies ONLY to new content
    const deselectedFormats = this.formatManager.getDeselectedFormats();
    const hasDeselectedFormats = Object.keys(deselectedFormats).length > 0;
    
    if (!isExistingContent && hasDeselectedFormats) {
      // Only filter for empty cells - respect user's deselection preferences
      console.log('🔍 updateToolbarUI: Empty cell - filtering deselected formats (GLOBAL preference)');
      filteredFormats = this.formatManager.filterDeselectedFormats(filteredFormats);
    } else if (isExistingContent) {
      // Existing content - NEVER filter, show actual formats
      console.log('🔍 updateToolbarUI: Existing content - NOT filtering deselected formats, showing actual cell formats');
    } else {
      console.log('🔍 updateToolbarUI: No deselected formats to filter');
    }
    
  // Update format buttons (bold, italic, underline, strike)
    ['bold', 'italic', 'underline', 'strike'].forEach(format => {
      const buttons = toolbarContainer.querySelectorAll(`.ql-${format}`);
      buttons.forEach(button => {
        if (filteredFormats[format]) {
          button.classList.add('ql-active');
          button.setAttribute('aria-pressed', 'true');
          console.log(`🔍 updateToolbarUI: Activated ${format} button`);
        } else {
          button.classList.remove('ql-active');
          button.setAttribute('aria-pressed', 'false');
          console.log(`🔍 updateToolbarUI: Deactivated ${format} button`);
        }
      });
    });

  
  // Update color buttons - only if not deselected (or we're in a non-empty cell)
  if (filteredFormats.color) {
    const colorButton = toolbarContainer.querySelector('.ql-color .ql-picker-label');
    if (colorButton) {
      const colorIcon = colorButton.querySelector('.ql-stroke');
      if (colorIcon) {
        (colorIcon as HTMLElement).style.stroke = filteredFormats.color;
      }
    }
  }

  // Update background color - only if not deselected (or we're in a non-empty cell)
  if (filteredFormats.background) {
    const bgButton = toolbarContainer.querySelector('.ql-background .ql-picker-label');
    if (bgButton) {
      const bgIcon = bgButton.querySelector('.ql-fill');
      if (bgIcon) {
        (bgIcon as HTMLElement).style.fill = filteredFormats.background;
      }
    }
  }

  // Update alignment buttons
  if (filteredFormats.align) {
    const alignButtons = document.querySelectorAll('.ql-align');
    alignButtons.forEach(button => {
      button.classList.remove('ql-active');
    });

    const activeAlignButton = document.querySelector(`.ql-align[value="${filteredFormats.align}"]`);
    if (activeAlignButton) {
      activeAlignButton.classList.add('ql-active');
    }
  }
  
  // Fix font/size pickers - set BOTH data-value AND data-label for proper display
  ['font', 'size'].forEach(format => {
    if (!filteredFormats[format]) return;
    
    const handler = toolbarContainer.querySelector(`.ql-${format}`);
    if (!handler) return;
    
    const label = handler.querySelector('.ql-picker-label');
    if (label) {
      label.setAttribute('data-value', filteredFormats[format]);
      label.setAttribute('data-label', format == 'size' ? `${filteredFormats[format]}s` : filteredFormats[format]); // ← CSS reads this for display
    }
    
    handler.querySelectorAll('.ql-picker-item').forEach(item => {
      item.classList.toggle('ql-selected',
        (item.getAttribute('data-value') || '').toLowerCase() === (filteredFormats[format] || '').toLowerCase()
      );
    });
  });
}

  /**
   * Helper methods that delegate to FormatManager
   * These maintain backward compatibility with existing code
   */
  
  getActiveFormats() {
    return this.formatManager.getGlobalFormats();
  }

  getDeselectedFormats() {
    return this.formatManager.getDeselectedFormats();
  }

  clearAllDeselectedFormats() {
    this.formatManager.clearDeselectedFormats();
  }

  clearActiveFormats() {
    this.formatManager.clearGlobalFormats();
  }


  // Get text formatting from previous cell
  getPreviousCellFormats(currentCell: any): Record<string, any> {
    try {
      if (!currentCell) {
        return {};
      }
      
      // Find the table containing this cell
      const table = currentCell.closest('table');
      if (!table) {
        return {};
      }
      
      // Get all cells in the table
      const cells = Array.from(table.querySelectorAll('td, th'));
      const currentIndex = cells.indexOf(currentCell);
      
      if (currentIndex === -1) {
        return {};
      }
      
      // Find the previous non-empty cell
      for (let i = currentIndex - 1; i >= 0; i--) {
        const prevCell: any = cells[i];
        const cellText = (prevCell as HTMLElement).innerText || '';
        const hasContent = cellText.replace(/\u200B/g, '').trim().length > 0;
        
        if (hasContent) {
          // Found a previous cell with content, get its formats
          // Skip filtering to detect all formats for inheritance
          const prevCellBlot = Quill.find(prevCell) as any;
          if (prevCellBlot) {
            const formats = this.getCellTextFormats(prevCellBlot, true); // Get all formats first
            // CRITICAL: Filter out deselected formats before returning
            const filteredFormats = this.formatManager.filterDeselectedFormats(formats);
            // Update lastAppliedFormats with FILTERED formats
            this.updateLastAppliedFormats(filteredFormats);
            
            return filteredFormats;
          }
        }
      }
      
      // No previous cell with content found, try to use lastAppliedFormats
      if (Object.keys(this.lastAppliedFormats).length > 0) {
        return { ...this.lastAppliedFormats };
      }
      return {};
      
    } catch (e) {
      console.error('🔍 getPreviousCellFormats: Error getting previous cell formats:', e);
      return {};
    }
  }

  // Track the last formats that were actually applied by user typing
  private lastAppliedFormats: any = {};

  // Update the last applied formats when user types with certain formats
  updateLastAppliedFormats(formats: any) {
    this.lastAppliedFormats = { ...formats };
  }

  // Get the last applied formats (for inheritance by next cells)
  getLastAppliedFormats(): any {
    return { ...this.lastAppliedFormats };
  }


  // Get text formatting from cell content (bold, italic, font, size, color, etc.)
  getCellTextFormats(cell: any, skipDeselectedFilter: boolean = false): any {
    if (!cell) return {};
    
    try {
      // Validate cell is a proper blot with length method
      if (typeof cell.length !== 'function') {
        console.log('🔍 getCellTextFormats: Cell is not a valid blot');
        return {};
      }
      
       // Use reliable text content check instead of cellLength > 1
      const cellText = (cell.domNode as HTMLElement).innerText || '';
      const hasContent = cellText.replace(/\u200B/g, '').trim().length > 0;
      const cellIndex = this.quill.getIndex(cell);
      
      if (cellIndex === null || cellIndex === undefined || cellIndex < 0) {
        console.log('🔍 getCellTextFormats: Invalid cellIndex, returning empty formats');
        return {};
      }
      
      // Get formats from the cell content
      // For empty cells, get formats from the cell start
      // For non-empty cells, get formats from the first character to represent current state
      let formats: any = {};
      if (hasContent) {
        // Cell has content: Get formats from the FIRST character of actual content
        // CRITICAL: Never read from the last character (cell boundary/newline) as it carries no font info
        // The first character is where the actual text formatting is applied
        const firstCharIndex = cellIndex;
        formats = this.quill.getFormat(firstCharIndex, 1);
      } else {
        // Empty cell: Get formats from the cell start position
        // This should include font and size if they were applied to the cell
        formats = this.quill.getFormat(cellIndex, 1);
        // If no formats at cell start, try to get them from activeFormats
        if (Object.keys(formats).length === 0) {
          const activeFormats = this.getActiveFormats();
          console.log("ActiveFormats", activeFormats);
          formats = { ...activeFormats };
          console.log('🔍 getCellTextFormats: Empty cell - using activeFormats as fallback:', formats);
        }
      }
      // CRITICAL: Filter out deselected formats from cell content
      // IMPORTANT: Only filter inline formats (bold, italic, underline, strike)
      // Font and size are structural formats that should ALWAYS persist
      // BUT skip filtering if we're detecting formats for inheritance
      if (!skipDeselectedFilter) {
        const deselectedFormats = this.formatManager.getDeselectedFormats();
        Object.keys(deselectedFormats).forEach(format => {
          if (deselectedFormats[format] && 
              formats[format] && 
              this.formatManager.isInlineFormat(format)) {
            delete formats[format];
            console.log(` getCellTextFormats: Filtered out deselected inline format ${format} from cell content`);
          }
        });
      } else {
        console.log(' getCellTextFormats: Skipping deselected format filter for inheritance detection');
      }
      return formats;
    } catch (e) {
      console.error(' getCellTextFormats: Error getting cell formats:', e);
      return {};
    }
  }

  /**
   * Remove formats from deselected list if they're already active in the current content
   * This is called when clicking on cells that have active formats
   * If a format is already applied to the cell content, it shouldn't be in the deselected list
   */
  removeDeselectedFormatsIfActive(cellFormats: any) {
    if (!cellFormats || Object.keys(cellFormats).length === 0) {
      console.log(' removeDeselectedFormatsIfActive: No cell formats provided');
      return;
    }

    const deselectedFormats = this.formatManager.getDeselectedFormats();
    
    // Check each deselected format
    Object.keys(deselectedFormats).forEach(format => {
      if (deselectedFormats[format] && this.formatManager.isInlineFormat(format)) {
        // If this inline format is active in the cell content, remove it from deselected list
        if (cellFormats[format]) {
          console.log(` removeDeselectedFormatsIfActive: Format "${format}" is active in cell content - removing from deselected`);
        }
      }
    });

    // Remove the formats by toggling them back on
    Object.keys(deselectedFormats).forEach((format: string) => {
      if (deselectedFormats[format] && this.formatManager.isInlineFormat(format) && cellFormats[format]) {
        this.formatManager.toggleGlobalFormat(format, cellFormats[format]);
      }
    });

    if (Object.keys(deselectedFormats).length > 0) {
      console.log(` removeDeselectedFormatsIfActive: Removed formats from deselected:`, Object.keys(deselectedFormats));
    }
  }

  handleKeyup(e: KeyboardEvent) {
    if (!this.quill.isEnabled()) return;
    this.cellSelection.handleKeyup(e);
    if (e.ctrlKey && (e.key === 'z' || e.key === 'y')) {
      this.hideTools();
      this.clearHistorySelected();
    } else {
      // Keep menu visible while typing in table cells
      const [table] = this.getTable();
      if (table) {
        this.tableMenus.showMenus();
      }
    }
    this.updateMenus(e);
  }

  handleMousedown(e: MouseEvent) {
    if (!this.quill.isEnabled()) return;
    this.tableSelect?.hide(this.tableSelect.root);
    const table = (e.target as Element).closest('table');
    
    // Check if click is immediately after a table
    const clickTarget = e.target as Element;
    const isParagraphAfterTable = !table && 
      clickTarget.nodeName === 'P' && 
      clickTarget.previousElementSibling?.nodeName === 'TABLE';
    
    // In-table Editor
    if (table && !this.quill.root.contains(table)) {
      this.hideTools();
      return;
    }
    
    if (!table) {
      // Check if we clicked on a cell (might be in a different table)
      const cell = clickTarget.closest('td, th');
      if (!cell) {
        this.hideTools();
      }
      
      // If clicking right after a table, prevent default behavior that might create a row
      if (isParagraphAfterTable) {
        // Just set selection without triggering row creation
        const range = this.quill.getSelection();
        if (range) {
          this.quill.setSelection(range.index, 0, Quill.sources.USER);
        }
        return;
      }
      
      if (!cell) {
        this.handleMouseMove();
      }
      return;
    }
    this.cellSelection.handleMousedown(e);
    this.cellSelection.setDisabled(true);
  }

  // If the default selection includes table cells,
  // automatically select the entire table
  handleMouseMove() {
    let table: Element = null;
    const handleMouseMove = (e: MouseEvent) => {
      if (!table) table = (e.target as Element).closest('table');
    }

    const handleMouseup = (e: MouseEvent) => {
      if (table) {
        const tableBlot = Quill.find(table);
        if (!tableBlot) return;
        // @ts-expect-error
        const index = tableBlot.offset(this.quill.scroll);
        // @ts-expect-error
        const length = tableBlot.length();
        const range = this.quill.getSelection();
        const minIndex = Math.min(range.index, index);
        const maxIndex = Math.max(range.index + range.length, index + length);
        this.quill.setSelection(
          minIndex,
          maxIndex - minIndex,
          Quill.sources.USER
        );
      } else {
        // Clicked outside table - update toolbar with current formats
        const range = this.quill.getSelection();
        if (range) {
          const currentFormats = this.quill.getFormat(range);
          this.updateToolbarUI(currentFormats);
        }
      }
      this.quill.root.removeEventListener('mousemove', handleMouseMove);
      this.quill.root.removeEventListener('mouseup', handleMouseup);
    }

    this.quill.root.addEventListener('mousemove', handleMouseMove);
    this.quill.root.addEventListener('mouseup', handleMouseup);
  }

  handleScroll() {
    if (!this.quill.isEnabled()) return;
    
    // Don't hide tools if a dropdown is currently open
    if (this.tableMenus?.getDropdownOpen()) {
      return;
    }
    
    this.hideTools();
    this.tableMenus?.updateScroll(true);
  }

  hideTools() {
    this.cellSelection?.clearSelected();
    this.cellSelection?.setDisabled(false);
    this.operateLine?.hideDragBlock();
    this.operateLine?.hideDragTable();
    this.operateLine?.hideLine();
    this.tableMenus?.hideMenus();
    this.tableMenus?.destroyTablePropertiesForm();
  }

  insertTable(rows: number, columns: number) {
    // Clear deselected formats for new table - each table should have fresh format state
    this.clearAllDeselectedFormats();
    this.clearActiveFormats();
    const range = this.quill.getSelection(true);
    if (range == null) return;
    if (this.isTable(range)) return;
    //const style = `width: 100%`;
    const formats = this.quill.getFormat(range.index - 1);
    const [, offset] = this.quill.getLine(range.index);
    const isExtra = !!formats[TableCellBlock.blotName] || offset !== 0;
    const _offset = isExtra ? 2 : 1;
    const extraDelta = isExtra ? new Delta().insert('\n') : new Delta();
    const base = new Delta()
      .retain(range.index)
      .delete(range.length)
      .concat(extraDelta)
      .insert('\n');
    const delta = new Array(rows).fill(0).reduce(memo => {
      const id = tableId();
      return new Array(columns).fill('\n').reduce((memo, text) => {
        return memo.insert(text, {
          [TableCellBlock.blotName]: cellId(),
          [TableCell.blotName]: { 'data-row': id }
        });
      }, memo);
    }, base);
    
    // // Store formats from previous line to apply to first cell of new table
    // // Filter out table-related formats and keep only text formatting
    // const previousLineFormats: any = {};
    // this.formatsToPersist.forEach(key => {
    //   if (formats[key] !== undefined && formats[key] !== false) {
    //     previousLineFormats[key] = formats[key];
    //   }
    // });
    
    // // Store these formats to be applied when first cell is clicked
    // (window as any).formatsFromPreviousLine = previousLineFormats;
    // console.log('🔍 TABLE INSERT: Stored formats from previous line:', previousLineFormats);
    
    // // Also immediately update activeFormats so they're available when cell is focused
    // Object.keys(previousLineFormats).forEach(key => {
    //   this.activeFormats[key] = previousLineFormats[key];
    // });
    // console.log('🔍 TABLE INSERT: Updated activeFormats with previous line formats:', this.activeFormats);
    
    // Mark that a new table is being created BEFORE showTools to prevent layout interference
    if (this.tableMenus) {
      this.tableMenus.markTableAsNewlyCreated();
    }
    // Clear any old cell selections to prevent scrolling to old tables
    if (this.cellSelection) {
      this.cellSelection.clearSelected();
    }
    this.quill.updateContents(delta, Quill.sources.USER);
    this.quill.setSelection(range.index + _offset, Quill.sources.SILENT);
    this.showTools();
  }

  // Inserting tables within tables is currently not supported
  private isTable(range: Range) {
    const formats = this.quill.getFormat(range.index);
    return !!formats[TableCellBlock.blotName];
  }

  // Completely delete empty tables
  listenDeleteTable() {
    this.quill.on(Quill.events.TEXT_CHANGE, (delta: any  , old: any, source: any ) => {
      if (source !== Quill.sources.USER) return;
      const tables = this.quill.scroll.descendants(TableContainer);
      if (!tables.length) return;
      const deleteTables: TableContainer[] = [];
      tables.forEach((table: any) => {
        const tbody = table.tbody();
        const thead = table.thead();
        if (!tbody && !thead) deleteTables.push(table);
      });
      if (deleteTables.length) {
        for (const table of deleteTables) {
          table.remove();
        }
        this.hideTools();
        this.quill.update(Quill.sources.API);
      }
    });
  }

  private registerToolbarTable(toolbarTable: boolean) {
    if (!toolbarTable) return;
    Quill.register({ 'formats/table-better': ToolbarTable }, true);
    const toolbar = this.quill.getModule('toolbar') as TableToolbar;
    const button = toolbar.container.querySelector('button.ql-table-better');
    if (!button || !this.tableSelect.root) return;
    button.appendChild(this.tableSelect.root);
    button.addEventListener('click', (e: MouseEvent) => {
      this.tableSelect.handleClick(e, this.insertTable.bind(this));
    });
    document.addEventListener('click', (e: MouseEvent) => {
      const visible = e.composedPath().includes(button);
      if (visible) return;
      if (!this.tableSelect.root.classList.contains('ql-hidden')) {
        this.tableSelect.hide(this.tableSelect.root);
      }
    });
  }

    showTools(force?: boolean) {
    const [table, , cell] = this.getTable();
    if (!table || !cell) return;
    
    try {
      this.cellSelection.setDisabled(true);
      
      // Check if table was recently created to prevent unwanted scrolling
      const timeSinceCreation = Date.now() - (this.tableMenus?.lastTableCreationTime || 0);
      const isRecentlyCreated = timeSinceCreation < 1000; // Within 1 second
      
      // Add null check before accessing cell.domNode
      if (cell && cell.domNode) {
        // Don't force selection (which triggers scrolling) for newly created tables
        // This prevents jumping to other tables when creating a new one
        const shouldForce = isRecentlyCreated ? false : (force !== false);
        this.cellSelection.setSelected(cell.domNode, shouldForce);
      }
      
      this.tableMenus.showMenus();
      
      // Add null checks before accessing table.domNode
      if (table && table.domNode) {
        this.tableMenus.updateMenus(table.domNode);
        this.tableMenus.updateTable(table.domNode);
      }
    } catch (error) {
      console.error('Error in showTools:', error);
    }
  }

  private updateMenus(e: KeyboardEvent) {
    if (!this.cellSelection.selectedTds.length) return;
    if (
      e.key === 'Enter' ||
      (e.ctrlKey && e.key === 'v')
    ) {
      this.tableMenus.updateMenus();
    }
  }
}

const keyboardBindings = {
  'table-cell down': makeTableArrowHandler(false),
  'table-cell up': makeTableArrowHandler(true),
  'table-cell-block backspace': makeCellBlockHandler('Backspace'),
  'table-cell-block delete': makeCellBlockHandler('Delete'),
  'table-header backspace': makeTableHeaderHandler('Backspace'),
  'table-header delete': makeTableHeaderHandler('Delete'),
  'table-header enter': {
    key: 'Enter',
    collapsed: true,
    format: ['table-header'],
    suffix: /^$/,
    handler(range: Range, context: Context) {
      const [line, offset] = this.quill.getLine(range.index);
      const delta = new Delta()
        .retain(range.index)
        .insert('\n', context.format)
        .retain(line.length() - offset - 1)
        .retain(1, { header: null });
      this.quill.updateContents(delta, Quill.sources.USER);
      this.quill.setSelection(range.index + 1, Quill.sources.SILENT);
      this.quill.scrollSelectionIntoView();
    },
  },
  'table-list backspace': makeTableListHandler('Backspace'),
  'table-list delete': makeTableListHandler('Delete'),
  'table-list empty enter': {
    key: 'Enter',
    collapsed: true,
    format: ['table-list'],
    empty: true,
    handler(range: Range, context: Context) {
      const { line } = context;
      const { cellId } = line.parent.formats()[line.parent.statics.blotName];
      const blot = line.replaceWith(TableCellBlock.blotName, cellId) as TableCellBlock;
      const tableModule = this.quill.getModule('table-better');
      const cell = getCorrectCellBlot(blot);
      cell && tableModule.cellSelection.setSelected(cell.domNode, false);
    }
  }
}

function makeCellBlockHandler(key: string) {
  return {
    key,
    format: ['table-cell-block', 'table-th-block'],
    collapsed: true,
    handler(range: Range, context: Context) {
      const [line] = this.quill.getLine(range.index);
      const { offset, suffix } = context;
      console.log("offset", offset);
      
      
      // If at the beginning of a cell
      if (offset === 0) {
        // Get the table module
        const tableModule = this.quill.getModule('table');
        if (!tableModule) return true;
        
        // Check if this is the first cell in the table
        const [table] = tableModule.getTable(range);
        if (table) {
          const tableIndex = this.quill.getIndex(table);
          
          // If we're at the beginning of the table, prevent backspace
          // This prevents content from above being pulled into the table
          if (range.index === tableIndex) {
            return false;
          }

          console.log("tableIndex", tableIndex);
          console.log("range.index", range.index);
        }
        
        if (!line.prev) return false;
        const blotName = line.prev?.statics.blotName;
        if (
          blotName === ListContainer.blotName ||
          blotName === TableCellBlock.blotName ||
          blotName === TableHeader.blotName
        ) {
          return removeLine.call(this, line, range);
        }
      }
      
      // Delete isn't from the end
      if (offset !== 0 && !suffix && key === 'Delete') {
        return false;
      }
      
      return true;
    }
  }
}
// Prevent table default up and down keyboard events.
// Implemented by the makeTableArrowVerticalHandler function.
function makeTableArrowHandler(up: boolean) {
  return {
    key: up ? 'ArrowUp' : 'ArrowDown',
    collapsed: true,
    format: ['table-cell', 'table-th'],
    handler() {
      return false;
    }
  };
}

function makeTableHeaderHandler(key: string) {
  return {
    key,
    format: ['table-header'],
    collapsed: true,
    empty: true,
    handler(range: Range, context: Context) {
      const [line] = this.quill.getLine(range.index);
      if (line.prev) {
        return removeLine.call(this, line, range);
      } else {
        const cellId = getCellId(line.formats()[line.statics.blotName]);
        line.replaceWith(TableCellBlock.blotName, cellId);
      }
    }
  }
}

function makeTableListHandler(key: string) {
  return {
    key,
    format: ['table-list'],
    collapsed: true,
    empty: true,
    handler(range: Range, context: Context) {
      const [line] = this.quill.getLine(range.index);
      const cellId = getCellId(line.parent.formats()[line.parent.statics.blotName]);
      line.replaceWith(TableCellBlock.blotName, cellId);      
    }
  }
}

function removeLine(line: Line, range: Range) {
  const tableModule = this.quill.getModule('table-better');
  line.remove();
  tableModule?.tableMenus.updateMenus();
  this.quill.setSelection(range.index - 1, Quill.sources.SILENT);
  return false;
}


Table.keyboardBindings = keyboardBindings;

export default Table;