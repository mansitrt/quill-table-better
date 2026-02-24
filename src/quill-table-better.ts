import Quill from 'quill';
import Delta from 'quill-delta';
import type { EmitterSource, Range } from 'quill';
import type { Props } from './types';
import type { BindingObject, Context } from './types/keyboard';
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
  lastFormat: any = null;
  defaultFormats: any = null;
  // Class field for tracking last clicked cell
  lastClickedCell: HTMLElement | null = null;
  private userDeselectedFormats: any = {}; // NEW: Track user deselections
  private activeFormats: any = {}; // NEW: Track active formats in table cells
  formatsToPersist = [
    'bold', 'italic', 'underline', 'strike',
    'font', 'size',
    'color', 'background',
    'align'
  ];
  isPerformingTableOperation = false;
  isSettingUpTableCell = false;

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
    this.defaultFormats = options.defaultFormats || {};
        
    // Override toolbar update method to prevent clearing ql-active classes during table cell selection
    const toolbar = this.quill.getModule('toolbar');
    if (toolbar && toolbar.update) {
      const originalUpdate = toolbar.update;
      toolbar.update = (range: any) => {
        // Skip toolbar updates when we're setting up a table cell to prevent clearing ql-active classes
        if (this.isSettingUpTableCell) {
          console.log('🔍 SKIPPING toolbar update during table cell setup');
          return;
        }
        // Call original update method
        originalUpdate.call(toolbar, range);
      };
    }
        
    // Add click handler to track the last clicked cell
    this.quill.root.addEventListener('click', (e: any) => {
      // Note: Cell tracking is now handled in table-menus.ts handleClick method
      // This listener is kept for backward compatibility but doesn't override table-menus tracking
    });
    
    quill.root.addEventListener('keyup', this.handleKeyup.bind(this));
    quill.root.addEventListener('mousedown', this.handleMousedown.bind(this));
    quill.root.addEventListener('scroll', this.handleScroll.bind(this));
    this.listenDeleteTable();
    this.registerToolbarTable(options?.toolbarTable);
    this.setupFormatPersistence();
    this.setupTableCellFormatting();
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

  setupTableCellFormatting() {
    // DISABLED: This click listener was interfering with table cell focus and selection
    // The table-menus.ts handleClick method now handles all table cell interactions properly
    // Removing this prevents selection resets and focus loss on normal taps
    /*
    this.quill.root.addEventListener('click', (event: any) => {
      // Get current selection
      let range = this.quill.getSelection();
      if (!range) {
        // Try to get cell from click event
        const clickedElement = event.target as Element;
        const cell = clickedElement.closest('td, th');
        if (cell) {
          const cellBlot = Quill.find(cell);
          if (cellBlot) {
            const cellIndex = this.quill.getIndex(cellBlot);
            this.quill.setSelection(cellIndex, 0, Quill.sources.SILENT);
            range = this.quill.getSelection();
          }
        }
        if (!range) return;
      }

      const tableModule = this.quill.getModule('table-better');
      if (!tableModule) return;

      // Check if we're in a table cell
      const [__, _, cell] = tableModule.getTable(range);
      if (!cell) return;

      // // Get the cell's text length to check if it has content
      // const cellIndex = this.quill.getIndex(cell);
      // const cellLength = cell.length();
      
      // // Get formats from the cell content (not just the cursor position)
      // let formats: any = {};
      // if (cellLength > 1) {
      //   // Cell has content - get formats from the first character of content
      //   formats = this.quill.getFormat(cellIndex, 1);
      // } else {
      //   // Empty cell - get formats at cursor
      //   formats = this.quill.getFormat(range);
      // }

      // console.log('Cell formats detected:', formats);

      // // Update toolbar UI immediately
      // setTimeout(() => {
      //   this.updateToolbarUI(formats);
      // }, 10);
    });
    */
  }

updateToolbarUI(formats: any) {
  const toolbarContainer = document.querySelector('.ql-toolbar');
  if (!toolbarContainer) return;

  // Filter out formats that user explicitly deselected
  const filteredFormats = { ...formats };
  Object.keys(this.userDeselectedFormats).forEach(format => {
    if (this.userDeselectedFormats[format]) {
      delete filteredFormats[format];
    }
  });

  // Update format buttons (bold, italic, underline, strike)
  ['bold', 'italic', 'underline', 'strike'].forEach(format => {
    const buttons = toolbarContainer.querySelectorAll(`.ql-${format}`);
    buttons.forEach(button => {
      if (filteredFormats[format]) {
        button.classList.add('ql-active');
        button.setAttribute('aria-pressed', 'true');
      } else {
        button.classList.remove('ql-active');
        button.setAttribute('aria-pressed', 'false');
      }
    });
  });

  // Update font picker using Quill's toolbar update mechanism
  if (formats.font) {
    const fontPicker = toolbarContainer.querySelector('.ql-font .ql-picker-label');
    if (fontPicker) {
      fontPicker.setAttribute('data-value', formats.font);
      fontPicker.classList.add('ql-active');
      // Update the visible text
      const labelText = fontPicker.querySelector('.ql-picker-label-text');
      if (labelText) {
        labelText.textContent = formats.font;
      }
    }

  } else {
    const fontPicker = toolbarContainer.querySelector('.ql-font .ql-picker-label');
    if (fontPicker) {
      fontPicker.classList.remove('ql-active');
    }
  }

  // Update size picker - both data attribute AND visible label
  if (formats.size) {
    const sizePicker = toolbarContainer.querySelector('.ql-size .ql-picker-label');
    if (sizePicker) {
      sizePicker.setAttribute('data-value', formats.size);
      sizePicker.classList.add('ql-active');
    }
  } else {
    // If no size format, remove active class
    const sizePicker = toolbarContainer.querySelector('.ql-size .ql-picker-label');
    if (sizePicker) {
      sizePicker.classList.remove('ql-active');
    }
  }

  // Update color buttons
  if (formats.color) {
    const colorButton = toolbarContainer.querySelector('.ql-color .ql-picker-label');
    if (colorButton) {
      const colorIcon = colorButton.querySelector('.ql-stroke');
      if (colorIcon) {
        (colorIcon as HTMLElement).style.stroke = formats.color;
      }
    }
  }

  if (formats.background) {
    const bgButton = toolbarContainer.querySelector('.ql-background .ql-picker-label');
    if (bgButton) {
      const bgIcon = bgButton.querySelector('.ql-fill');
      if (bgIcon) {
        (bgIcon as HTMLElement).style.fill = formats.background;
      }
    }
  }
}



  setupFormatPersistence() {
    let isInTable = false;
    let currentCell: any = null;
    // activeFormats is now a class property, no need to redeclare
    const activeFormats = this.activeFormats;
    let isPerformingTableOperation = false;

    // List of formats we want to persist
    const formatsToPersist = [
      'bold', 'italic', 'underline', 'strike',
      'font', 'size',
      'color', 'background',
      'align'
    ];

    const updateActiveFormats = (range: any) => {
      if (!range) return;
      const formats = this.quill.getFormat(range);

      // Update active formats, focusing on the ones we want to persist
      formatsToPersist.forEach(key => {
        if (formats[key] !== undefined) {
          activeFormats[key] = formats[key];
        }
      });
    };

    // Listen for toolbar format changes to detect when user deselects formats
    const updateFormatsFromToolbar = () => {
      if (!isInTable || !currentCell) return;
      
      const range = this.quill.getSelection();
      if (!range) return;
      
      // Check toolbar button states directly instead of cursor formats
      const toolbarContainer = toolbar?.container;
      if (!toolbarContainer) return;
      
      // Update activeFormats based on toolbar button states
      ['bold', 'italic', 'underline', 'strike'].forEach(format => {
        const button = toolbarContainer.querySelector(`.ql-${format}`);
        if (button) {
          const isActive = button.classList.contains('ql-active');
          const wasActive = activeFormats[format] === true;
          
          if (isActive) {
            activeFormats[format] = true;
            // User turned it on - remove from deselected list
            delete this.userDeselectedFormats[format];
          } else {
            // Button is not active
            if (wasActive) {
              // User just turned it off - mark as explicitly deselected
              this.userDeselectedFormats[format] = true;
              console.log(`User deselected ${format}`);
            }
            delete activeFormats[format];
          }
        }
      });
      
      // For other formats (font, size, color, etc.), check their values
      const currentFormats = this.quill.getFormat(range);
      ['font', 'size', 'color', 'background', 'align'].forEach(format => {
        if (currentFormats[format] !== undefined && currentFormats[format] !== false) {
          activeFormats[format] = currentFormats[format];
        } else if (!currentFormats[format]) {
          delete activeFormats[format];
        }
      });
      
      console.log('Updated activeFormats from toolbar:', activeFormats);
      console.log('User deselected formats:', this.userDeselectedFormats);
    };

    // Listen for toolbar clicks
    const toolbar = this.quill.getModule('toolbar');
    if (toolbar && toolbar.container) {
      toolbar.container.addEventListener('click', (e: MouseEvent) => {
        // Small delay to let Quill process the format change first
        setTimeout(() => {
          updateFormatsFromToolbar();
        }, 50);
      });
      
      // Also listen for select changes (for font, size, etc.)
      toolbar.container.addEventListener('change', (e: Event) => {
        setTimeout(() => {
          updateFormatsFromToolbar();
        }, 50);
      });
    }

    // Handle selection changes - MINIMAL PROCESSING to prevent table mutations
    this.quill.on('selection-change', (range: any, oldRange: any, source: string) => {
      // Skip SILENT selections to prevent unwanted mutations
      if (source === Quill.sources.SILENT) {
        return;
      }

      // REMOVED: Guard for table cell setup - we now use SILENT source so no selection-change events are triggered

      // Only track table state, don't trigger any updates
      if (!range) {
        isInTable = false;
        currentCell = null;
        return;
      }

      const tableModule = this.quill.getModule('table-better');
      if (!tableModule) return;

      const [__, _, cell] = tableModule.getTable(range);

      if (!cell) {
        isInTable = false;
        currentCell = null;
        return;
      }

      // Just track that we're in a table - no format processing or updates for SILENT source
      isInTable = true;
      currentCell = cell;
    });

    // Track format changes AND apply them to new text
    this.quill.on('text-change', (delta: any, oldContents: any, source: string) => {
      if (!isInTable || !currentCell) return;

      // Keep menu visible during all editor changes
      if (this.tableMenus && this.cellSelection.selectedTds.length > 0) {
        this.tableMenus.showMenus();
      }

      const range = this.quill.getSelection();
      if (!range) return;

      if (source === Quill.sources.USER) {
        // Remove any formats that user explicitly deselected
        const formatsToApply = { ...activeFormats };
        Object.keys(this.userDeselectedFormats).forEach(format => {
          delete formatsToApply[format];
        });
        
        // Apply formats to newly typed text
        if (Object.keys(formatsToApply).length > 0) {
          let position = 0;
          for (const op of delta.ops) {
            if (op.retain) {
              position += op.retain;
            } else if (op.insert && typeof op.insert === 'string') {
              // Apply formats to the inserted text using formatText (safe for tables)
              this.quill.formatText(position, op.insert.length, formatsToApply, Quill.sources.SILENT);
              // Update toolbar UI to show active formats
              this.updateToolbarUI(formatsToApply);
              break;
            }
          }
        }
        
        // Update toolbar to show current formats at cursor position
        const currentFormats = this.quill.getFormat(range);
        if (Object.keys(currentFormats).length > 0) {
          this.updateToolbarUI(currentFormats);
        }
        
        setTimeout(() => {
          // Delay format update to ensure all changes are processed
          updateActiveFormats(range);
        }, 10);
      }
    });

    // Handle format changes
    this.quill.on('editor-change', (eventName: string, ...args: any[]) => {
      if (!isInTable || !currentCell) return;

      // Keep menu visible during all editor changes
      if (this.tableMenus && this.cellSelection.selectedTds.length > 0) {
        this.tableMenus.showMenus();
      }

      const range = this.quill.getSelection();
      if (!range) return;

      if (eventName === 'text-change' || eventName === 'selection-change') {
        const formats = this.quill.getFormat(range);
        if (Object.keys(formats).length > 0) {
          formatsToPersist.forEach(key => {
            if (formats[key] !== undefined) {
                 // Only add the format if user hasn't explicitly deselected it
                if (!this.userDeselectedFormats[key]) {
                  activeFormats[key] = formats[key];
                }
            }
          });
        }
      }
    });

    // Hook into table operations
    const originalInsertColumn = this.tableMenus.insertColumn;
    this.tableMenus.insertColumn = function (...args: [HTMLTableColElement, number]) {
      isPerformingTableOperation = true;
      
      const capturedFormats = { ...activeFormats };
      
      const getUserDeselectedFormats = () => {
        try {
          if (typeof window !== 'undefined' && (window as any).userDeselectedFormats) {
            return (window as any).userDeselectedFormats;
          }
        } catch (e) {
          console.warn('Could not access userDeselectedFormats:', e);
        }
        return {};
      };

      const userDeselected = getUserDeselectedFormats();
      Object.keys(userDeselected).forEach(format => {
        if (userDeselected[format]) {
          delete capturedFormats[format];
        }
      });

      const result = originalInsertColumn.apply(this, args);
      
      setTimeout(() => {
        isPerformingTableOperation = false;
      }, 100);

      return result;
    };

    const originalInsertRow = this.tableMenus.insertRow;
    this.tableMenus.insertRow = function (...args: [HTMLElement, number]) {
      isPerformingTableOperation = true;
      
      const capturedFormats = { ...activeFormats };
      
      const getUserDeselectedFormats = () => {
        try {
          if (typeof window !== 'undefined' && (window as any).userDeselectedFormats) {
            return (window as any).userDeselectedFormats;
          }
        } catch (e) {
          console.warn('Could not access userDeselectedFormats:', e);
        }
        return {};
      };

      const userDeselected = getUserDeselectedFormats();
      Object.keys(userDeselected).forEach(format => {
        if (userDeselected[format]) {
          delete capturedFormats[format];
        }
      });

      const result = originalInsertRow.apply(this, args);
      
      setTimeout(() => {
        isPerformingTableOperation = false;
      }, 100);

      return result;
    };
  }

  // Expose activeFormats for table menu toolbar updates
  getActiveFormats() {
    return this.activeFormats || {};
  }

  safelyApplyFormats(cell: any, range: any, formats: any) {
  if (Object.keys(formats).length === 0 || this.isPerformingTableOperation) return;

  const cellIndex = this.quill.getIndex(cell);
  if (cellIndex == null || cellIndex === -1) return;

  // Filter formats to only include those we want to persist
  const filteredFormats: any = {};
  Object.keys(formats).forEach(key => {
    if (
      formats[key] !== undefined &&
      formats[key] !== false && 
      this.formatsToPersist.includes(key)
    ) {
      filteredFormats[key] = formats[key];
    }
  });

  if (Object.keys(filteredFormats).length === 0) return;

  console.log("Applying formats to cell:", filteredFormats);

  // Check if cell has content (excluding zero-width spaces)
  const cellText = cell.domNode?.innerText || '';
  const hasContent = cellText.replace(/\u200B/g, '').trim().length > 0;
  
  if (hasContent) {
    // Apply formats to existing content using formatText
    // Don't change selection - let user continue typing
    this.quill.formatText(cellIndex, cell.length() - 1, filteredFormats, Quill.sources.SILENT);
    
    // Update toolbar
    setTimeout(() => {
      if (typeof (window as any).updateToolbarUI === 'function') {
        (window as any).updateToolbarUI(filteredFormats);
      }
    }, 20);
  }
  // For empty cells, do nothing - let Quill handle formatting naturally
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