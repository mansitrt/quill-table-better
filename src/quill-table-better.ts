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
  formatsToPersist = [
    'bold', 'italic', 'underline', 'strike',
    'font', 'size',
    'color', 'background',
    'align'
  ];
  isPerformingTableOperation = false;

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
        
    // Add click handler to track the last clicked cell
    this.quill.root.addEventListener('click', (e: any) => {
      const target = e.target as HTMLElement;
      const cell = target.closest('td, th');
      if (cell) {
        this.lastClickedCell = cell as HTMLElement;
        console.log('Last clicked cell tracked:', cell);
      }
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
    this.quill.root.addEventListener('click', (event: any) => {


      console.log('event', event);
      // Get current selection
      const range = this.quill.getSelection();
      if (!range) return;

      console.log("range", range);

      const tableModule = this.quill.getModule('table-better');
      if (!tableModule) return true;

      // Check if we're in a table cell
      const [__, _, cell] = tableModule.getTable(range);
      if (!cell) return;

      // Get formats at current position
      let formats = this.quill.getFormat(range);

      // If the cell is empty, apply inherited or last-used formats
      if (this.cellIsEmpty(cell)) {

        console.log("coming here check it");
        // Try to get formats from non-empty cells above or fallback to last used
        let inheritedFormats = {};
        // 1. Try to get formats from the previous non-empty cell in the same column
        let row: any = cell.parent;
        let colIndex = Array.from(row.children).indexOf(cell);
        let prevRow = row.prev;
        while (prevRow) {
          let prevCell = prevRow?.children[colIndex];
          if (prevCell && !this.cellIsEmpty(prevCell)) {
            const prevCellIndex = this.quill.getIndex(prevCell);
            if (prevCellIndex != null && prevCellIndex !== -1) {
              inheritedFormats = this.quill.getFormat(prevCellIndex);
            }
          }
          prevRow = prevRow.prev;
        }
        // 2. If nothing found, fallback to lastFormat or defaultFormats
        if (!Object.keys(inheritedFormats).length) {
          inheritedFormats = this.lastFormat || this.defaultFormats || {};
        }
        setTimeout(() => {
          // Find the new cell and set selection
          const range = this.quill.getSelection();
          const tableModule = this.quill.getModule('table-better');
          if (!tableModule) return true;
          const [__, _, cell] = tableModule.getTable(range);
          if (cell) {
            // Apply fallback formats if needed
            this.safelyApplyFormats(cell, range, inheritedFormats);
            // --- Force toolbar update ---
            if (range) {
              // 1. Re-apply the selection to trigger Quill's internal update
              this.quill.setSelection(range.index, range.length, Quill.sources.SILENT);

              // 2. Emit a selection-change event to force all modules (including toolbar) to update
              if (this.quill.emitter) {
                this.quill.emitter.emit('selection-change', range, range, Quill.sources.USER);
              }
            }

          }
        }, 50);

      }

      // Find toolbar elements directly
      setTimeout(() => {

        // Find all toolbar buttons and selects
        const toolbarContainer = document.querySelector('.ql-toolbar');
        if (!toolbarContainer) return;

        // Process all format buttons
        Object.keys(formats).forEach(format => {
          const value = formats[format];

          // Handle buttons with specific values (like headers, list, etc.)
          if (value !== true) {
            // For buttons with values (like h1, h2, ordered list, etc.)
            const formatButtons = toolbarContainer.querySelectorAll(`.ql-${format}`);
            formatButtons.forEach(button => {
              const buttonValue = button.getAttribute('value');
              const isActive = buttonValue === value ||
                (buttonValue && value && buttonValue.toString() === value.toString());

              if (isActive) {
                button.classList.add('ql-active');
                button.setAttribute('aria-pressed', 'true');
              } else {
                button.classList.remove('ql-active');
                button.setAttribute('aria-pressed', 'false');
              }
            });

            // Handle select elements (like font, size, etc.)
            const selectElements = toolbarContainer.querySelectorAll(`select.ql-${format}`);
            selectElements.forEach(select => {
              const options = select.querySelectorAll('option');
              options.forEach(option => {
                if (option.value === value || option.value === value?.toString()) {
                  (option as HTMLOptionElement).selected = true;
                }
              });
            });
          } else {
            // For toggle buttons (like bold, italic, etc.)
            const buttons = toolbarContainer.querySelectorAll(`.ql-${format}`);
            buttons.forEach(button => {
              button.classList.add('ql-active');
              button.setAttribute('aria-pressed', 'true');
            });
          }
        });

        // Remove active class from buttons that don't match current formats
        const allButtons = toolbarContainer.querySelectorAll('button.ql-active');
        allButtons.forEach(button => {
          let className = Array.from(button.classList).find(c => c.startsWith('ql-'));
          if (!className) return;

          const format = className.substring(3); // Remove 'ql-' prefix
          if (!formats[format]) {
            button.classList.remove('ql-active');
            button.setAttribute('aria-pressed', 'false');
          }
        });
      }, 10);
    });
  }


  setupFormatPersistence() {
    let isInTable = false;
    let currentCell: any = null;
    let activeFormats: any = {};
    let isPerformingTableOperation = false; // Flag to track when table operations are happening

    // List of formats we want to persist
    const formatsToPersist = [
      'bold', 'italic', 'underline', 'strike',
      'font', 'size',
      'color', 'background',
      'align'
    ];

    // Helper method to safely apply formats without interfering with table operations
    const safelyApplyFormats = (cell: any, range: any, formats: any) => {
      if (Object.keys(formats).length === 0 || isPerformingTableOperation) return;

      // Get cell content length
      const length = cell?.length() || 0;

      // Safely apply formats
      const applyFormats = () => {
        if (isPerformingTableOperation) return; // Don't apply during table operations

        // Apply formats to cursor position
        Object.keys(formats).forEach(key => {
          if (formats[key] !== undefined && formatsToPersist.includes(key)) {
            try {
              this.quill.format(key, formats[key], Quill.sources.USER);
              if (length === 0) {
                this.quill.formatText(range.index, 1, { [key]: formats[key] }, Quill.sources.USER);
              }
            } catch (e) {
              console.log(`Error applying format ${key}:`, e);
            }
          }
        });
      };

      // Use setTimeout to ensure DOM is stable
      setTimeout(applyFormats, 10);
    };

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

    // Handle selection changes
    this.quill.on('selection-change', (range: any, oldRange: any, source: string) => {
      if (!range) {
        isInTable = false;
        currentCell = null;
        return;
      }

      const tableModule = this.quill.getModule('table-better');
      if (!tableModule) return true;

      const [__, _, cell] = tableModule.getTable(range);

      if (!cell) {
        isInTable = false;
        currentCell = null;
        return;
      }

      // Preserve existing formats before any changes
      if (oldRange) {
        const oldFormats = this.quill.getFormat(oldRange);
        Object.keys(oldFormats).forEach(key => {
          if (formatsToPersist.includes(key) && oldFormats[key] !== undefined) {
            activeFormats[key] = oldFormats[key];
          }
        });
      }

      // Handle cell change
      const handleCellChange = () => {
        const previousFormats = this.quill.getFormat(oldRange);
        if (!isInTable || currentCell !== cell) {
          if (!isInTable) {
            isInTable = true;
            // Get default formats only once when entering a table
            //    const defaultFormats : any = this.getDefaultFormats() || {};

            // Only apply default formats that we want to persist
            Object.keys(previousFormats).forEach(key => {
              if (formatsToPersist.includes(key) && previousFormats[key] !== undefined) {
                activeFormats[key] = previousFormats[key];
              }
            });


          }
          updateActiveFormats(range);
          safelyApplyFormats(cell, range, activeFormats);
          // if (!isPerformingTableOperation) {
          //   safelyApplyFormats(cell, range, activeFormats);
          // }
          currentCell = cell;
        }
      };

      // Use setTimeout for better stability
      setTimeout(handleCellChange, 10);
    });

    // Track format changes
    this.quill.on('text-change', (delta: any, oldContents: any, source: string) => {
      if (!isInTable || !currentCell) return;

      const range = this.quill.getSelection();
      if (!range) return;

      if (source === Quill.sources.USER) {
        setTimeout(() => {
          // Delay format update to ensure all changes are processed
          updateActiveFormats(range);
        }, 10);
      }
    });

    // Handle format changes
    this.quill.on('editor-change', (eventName: string, ...args: any[]) => {
      if (!isInTable || !currentCell) return;

      const range = this.quill.getSelection();
      if (!range) return;

      if (eventName === 'text-change' || eventName === 'selection-change') {
        const formats = this.quill.getFormat(range);
        if (Object.keys(formats).length > 0) {
          formatsToPersist.forEach(key => {
            if (formats[key] !== undefined) {
              activeFormats[key] = formats[key];
            }
          });
        }
      }
    });


    
  // Hook into other table operations similarly
  const originalInsertColumn = this.tableMenus.insertColumn;
  this.tableMenus.insertColumn = function (...args: [HTMLTableColElement, number]) {
    isPerformingTableOperation = true;
    const capturedFormats = { ...activeFormats };
    
    // Add a delay before opening dropdown menus on iOS to allow keyboard to close
    // This addresses the issue where dropdown menus don't position correctly when keyboard is open
    if (typeof window !== 'undefined' && /iPad|iPhone|iPod/.test(navigator.userAgent)) {
      // Store current window height to detect keyboard visibility
      const windowHeight = window.innerHeight;
      // If keyboard is likely visible (window height is reduced), delay the operation
      if (windowHeight < window.outerHeight * 0.8) {
        // Get the cell to use - either from args or the last clicked cell
        const [td, offset] = args;
        const tableModule = this.quill.getModule('table-better') as Table;
        const cellToUse = td?.isConnected ? td : tableModule.lastClickedCell as HTMLTableColElement;
        
        if (!cellToUse || !cellToUse.isConnected) {
          console.warn('No valid cell reference found for insertColumn operation');
          return;
        }
        
        // Wait for keyboard to close
        setTimeout(() => {
          try {
            const result = originalInsertColumn.call(this, cellToUse, offset);
            return result;
          } catch (error) {
            console.error('Error in delayed insertColumn operation:', error);
          }
        }, 300);
        return;
      }
    }

    const result = originalInsertColumn.apply(this, args);

    setTimeout(() => {
      isPerformingTableOperation = false;
      const range = this.quill.getSelection();
      if (range) {
        const tableModule = this.quill.getModule('table-better');
        if (tableModule) {
          const [table, _, cell] = tableModule.getTable(range);
          if (cell && table) {
            // For new cells, apply the formats that were active before the operation
            safelyApplyFormats(cell, range, capturedFormats);
          }
        }
      }
    }, 50); // Give it enough time for the DOM to stabilize

    return result;
    };
    
    // Hook into table operations to properly manage format preservation
    const originalInsertRow = this.tableMenus.insertRow;
    this.tableMenus.insertRow = function (...args) {
      isPerformingTableOperation = true;
      // Capture formats before the operation
      const capturedFormats = { ...activeFormats };

      // Add a delay before operations on iOS to allow keyboard to close
      if (typeof window !== 'undefined' && /iPad|iPhone|iPod/.test(navigator.userAgent)) {
        // Store current window height to detect keyboard visibility
        const windowHeight = window.innerHeight;
        // If keyboard is likely visible (window height is reduced), delay the operation
        if (windowHeight < window.outerHeight * 0.8) {
          // Get the cell to use - either from args or the last clicked cell
          const [td, offset] = args;
          const tableModule = this.quill.getModule('table-better') as Table;
          const cellToUse = td?.isConnected ? td : tableModule.lastClickedCell as HTMLTableColElement;
          
          if (!cellToUse || !cellToUse.isConnected) {
            console.warn('No valid cell reference found for insertRow operation');
            return;
          }
          
          // Wait for keyboard to close
          setTimeout(() => {
            try {
              const result = originalInsertRow.call(this, cellToUse, offset);
              return result;
            } catch (error) {
              console.error('Error in delayed insertRow operation:', error);
            }
          }, 300);
          return;
        }
      }

      // Perform the original operation
      const result = originalInsertRow.apply(this, args);

      // After the row is added, we'll apply the formats to the new cells
      setTimeout(() => {
        isPerformingTableOperation = false;
        // Get the new cells in the added row and apply formats
        const range = this.quill.getSelection();
        if (range) {
          const tableModule = this.quill.getModule('table-better');
          if (tableModule) {
            const [table, _, cell] = tableModule.getTable(range);
            if (cell && table) {
              // For new cells, apply the formats that were active before the operation
              safelyApplyFormats(cell, range, capturedFormats);
            }
          }
        }
      }, 50); // Give it enough time for the DOM to stabilize

      return result;
    };
  }

  safelyApplyFormats(cell: any, range: any, formats: any) {
    if (Object.keys(formats).length === 0 || this.isPerformingTableOperation) return;

    // const length = cell?.length() || 0;
    const applyFormats = () => {
      if (this.isPerformingTableOperation) return;
      const cellIndex = this.quill.getIndex(cell);
      if (cellIndex != null && cellIndex !== -1) {
        // Always remove all formats first
        this.quill.removeFormat(cellIndex, 1, Quill.sources.USER);

        // Only apply new formats if any are set to true/non-false
        const filteredFormats: any = {};
        Object.keys(formats).forEach(key => {
          if (
            formats[key] !== undefined &&
            formats[key] !== false && // Don't apply false (removal)
            this.formatsToPersist.includes(key)
          ) {
            filteredFormats[key] = formats[key];
          }
        });

        if (Object.keys(filteredFormats).length > 0) {
          this.quill.formatText(cellIndex, 1, filteredFormats, Quill.sources.USER);
        }
      }
    };
    setTimeout(applyFormats, 10);
  }

  handleKeyup(e: KeyboardEvent) {
    if (!this.quill.isEnabled()) return;
    this.cellSelection.handleKeyup(e);
    if (e.ctrlKey && (e.key === 'z' || e.key === 'y')) {
      this.hideTools();
      this.clearHistorySelected();
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
      this.hideTools();
      
      // If clicking right after a table, prevent default behavior that might create a row
      if (isParagraphAfterTable) {
        // Just set selection without triggering row creation
        const range = this.quill.getSelection();
        if (range) {
          this.quill.setSelection(range.index, 0, Quill.sources.USER);
        }
        return;
      }
      
      this.handleMouseMove();
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
      }
      this.quill.root.removeEventListener('mousemove', handleMouseMove);
      this.quill.root.removeEventListener('mouseup', handleMouseup);
    }

    this.quill.root.addEventListener('mousemove', handleMouseMove);
    this.quill.root.addEventListener('mouseup', handleMouseup);
  }

  handleScroll() {
    if (!this.quill.isEnabled()) return;
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
      
      // Add null check before accessing cell.domNode
      if (cell && cell.domNode) {
        this.cellSelection.setSelected(cell.domNode, force);
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