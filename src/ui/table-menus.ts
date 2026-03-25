import Quill from 'quill';
import Delta from 'quill-delta';
import merge from 'lodash.merge';
import type { LinkedList } from 'parchment';
import type {
  CorrectBound,
  Props,
  QuillTableBetter,
  TableCellMap,
  TableColgroup,
  TableContainer,
  UseLanguageHandler
} from '../types';
import {
  createTooltip,
  getAlign,
  getCellFormats,
  getCorrectBounds,
  getComputeBounds,
  getComputeSelectedCols,
  getComputeSelectedTds,
  setElementProperty,
  getElementStyle,
  updateTableWidth
} from '../utils';
import columnIcon from '../assets/icon/column.svg';
import rowIcon from '../assets/icon/row.svg';
import mergeIcon from '../assets/icon/merge.svg';
import tableIcon from '../assets/icon/table.svg';
import cellIcon from '../assets/icon/cell.svg';
import wrapIcon from '../assets/icon/wrap.svg';
import downIcon from '../assets/icon/down.svg';
import deleteIcon from '../assets/icon/delete.svg';
import copyIcon from '../assets/icon/copy.svg';
import {
  TableCell,
  tableId,
  TableTh,
  TableRow,
  TableThRow,
  TableThead
} from '../formats/table';
import TablePropertiesForm from './table-properties-form';
import {
  CELL_DEFAULT_VALUES,
  CELL_DEFAULT_WIDTH,
  CELL_PROPERTIES,
  DEVIATION,
  TABLE_PROPERTIES
} from '../config';
import Table from '../quill-table-better';

// Extend Window interface to include our custom property
declare global {
  interface Window {
    isSettingUpTableCell?: boolean;
    isProgrammaticallySettingSelection?: boolean;
  }
}

interface Children {
  [propName: string]: {
    content: string;
    handler: () => void;
    divider?: boolean;
    createSwitch?: boolean;
  }
}

interface Menu {
  content: string;
  icon: string;
  handler: (list: HTMLUListElement, tooltip: HTMLDivElement) => void;
  children?: Children;
}

interface CustomMenu extends Menu {
  name: 'column' | 'row' | 'wrap' | 'delete' | 'copy';
}

interface MenusDefaults {
  [propName: string]: Menu
}

enum Alignment {
  left = 'margin-left',
  right = 'margin-right'
}

function getMenusConfig(useLanguage: UseLanguageHandler, menus?: string[]): MenusDefaults {
  const DEFAULT: MenusDefaults = {
    column: {
      content: useLanguage('col'),
      icon: columnIcon,
      handler(list: HTMLUListElement, tooltip: HTMLDivElement) {
        this.toggleAttribute(list, tooltip);
      },
      children: {
        left: {
          content: useLanguage('insColL'),
          handler() {
            // Check if we're on iOS/iPad and use the helper method
            if (/iPad|iPhone|iPod/.test(navigator.userAgent)) {
              if (this.handleIOSTableOperation('column', 0)) {
                return;
              }
            }
            
            const selectedInfo = this.getSelectedTdsInfo();
            if (!selectedInfo) {
              console.error('Cannot insert column: no valid cell selected');
              return;
            }
            
            const { leftTd } = selectedInfo;
            if (!leftTd || !leftTd.isConnected) {
              console.error('Cannot insert column: leftTd is invalid or disconnected');
              return;
            }
            
            const bounds = this.table.getBoundingClientRect();
            this.insertColumn(leftTd, 0);
            updateTableWidth(this.table, bounds, CELL_DEFAULT_WIDTH);
            this.updateMenus();
          }
        },
        right: {
          content: useLanguage('insColR'),
          handler() {
            // Check if we're on iOS/iPad and use the helper method
            if (/iPad|iPhone|iPod/.test(navigator.userAgent)) {
              if (this.handleIOSTableOperation('column', 1)) {
                return;
              }
            }
            
            const selectedInfo = this.getSelectedTdsInfo();
            if (!selectedInfo) {
              console.error('Cannot insert column: no valid cell selected');
              return;
            }
            
            const { rightTd } = selectedInfo;
            if (!rightTd || !rightTd.isConnected) {
              console.error('Cannot insert column: rightTd is invalid or disconnected');
              return;
            }
            
            const bounds = this.table.getBoundingClientRect();
            this.insertColumn(rightTd, 1);
            updateTableWidth(this.table, bounds, CELL_DEFAULT_WIDTH);
            this.updateMenus();
          }
        },
        delete: {
          content: useLanguage('delCol'),
          handler() {
            this.deleteColumn();
          }
        }
        // select: {
        //   content: useLanguage('selCol'),
        //   handler() {
        //     this.selectColumn();
        //   }
        // }
      }
    },
    row: {
      content: useLanguage('row'),
      icon: rowIcon,
      handler(list: HTMLUListElement, tooltip: HTMLDivElement, e?: PointerEvent) {
        this.toggleAttribute(list, tooltip, e);
      },
      children: {
        // header: {
        //   content: useLanguage('headerRow'),
        //   divider: true,
        //   createSwitch: true,
        //   handler() {
        //     this.toggleHeaderRow();
        //     this.toggleHeaderRowSwitch();
        //   }
        // },
        above: {
          content: useLanguage('insRowAbv'),
          handler() {
            // Check if we're on iOS/iPad and use the helper method
            if (/iPad|iPhone|iPod/.test(navigator.userAgent)) {
              if (this.handleIOSTableOperation('row', 0)) {
                return;
              }
            }
            
            const selectedInfo = this.getSelectedTdsInfo();
            if (!selectedInfo) {
              console.error('Cannot insert row: no valid cell selected');
              return;
            }
            
            const { topTd } = selectedInfo;
            if (!topTd || !topTd.isConnected) {
              console.error('Cannot insert row: topTd is invalid or disconnected');
              return;
            }
            
            this.insertRow(topTd, -1);
            this.updateMenus();
          }
        },
        below: {
          content: useLanguage('insRowBlw'),
          handler() {
            // Check if we're on iOS/iPad and use the helper method
            if (/iPad|iPhone|iPod/.test(navigator.userAgent)) {
              if (this.handleIOSTableOperation('row', 1)) {
                return;
              }
            }
            
            const selectedInfo = this.getSelectedTdsInfo();
            if (!selectedInfo) {
              console.error('Cannot insert row: no valid cell selected');
              return;
            }
            
            const { bottomTd } = selectedInfo;
            if (!bottomTd || !bottomTd.isConnected) {
              console.error('Cannot insert row: bottomTd is invalid or disconnected');
              return;
            }
            
            this.insertRow(bottomTd, 1);
            this.updateMenus();
          }
        },
        delete: {
          content: useLanguage('delRow'),
          handler() {
            this.deleteRow();
          }
        }
        // select: {
        //   content: useLanguage('selRow'),
        //   handler() {
        //     this.selectRow();
        //   }
        // }
      }
    },
    // table: {
    //   content: useLanguage('tblProps'),
    //   icon: tableIcon,
    //   handler(list: HTMLUListElement, tooltip: HTMLDivElement) {
    //     const attribute = {
    //       ...getElementStyle(this.table, TABLE_PROPERTIES),
    //       'align': this.getTableAlignment(this.table)
    //     };
    //     this.toggleAttribute(list, tooltip);
    //     this.tablePropertiesForm = new TablePropertiesForm(this, { attribute, type: 'table' });
    //     this.hideMenus();
    //   }
    // },
    // cell: {
    //   content: useLanguage('cellProps'),
    //   icon: cellIcon,
    //   handler(list: HTMLUListElement, tooltip: HTMLDivElement) {
    //     const { selectedTds } = this.tableBetter.cellSelection;
    //     const attribute =
    //       selectedTds.length > 1
    //         ? this.getSelectedTdsAttrs(selectedTds)
    //         : this.getSelectedTdAttrs(selectedTds[0]);
    //     this.toggleAttribute(list, tooltip);
    //     this.tablePropertiesForm = new TablePropertiesForm(this, { attribute, type: 'cell' });
    //     this.hideMenus();
    //   }
    // },
    wrap: {
      content: useLanguage('insParaOTbl'),
      icon: wrapIcon,
      handler(list: HTMLUListElement, tooltip: HTMLDivElement) {
        this.toggleAttribute(list, tooltip);
      },
      children: {
        before: {
          content: useLanguage('insB4'),
          handler() {
            this.insertParagraph(-1);
          }
        },
        after: {
          content: useLanguage('insAft'),
          handler() {
            this.insertParagraph(1);
          }
        }
      }
    },
    delete: {
      content: useLanguage('delTable'),
      icon: deleteIcon,
      handler() {
        this.deleteTable();
      }
    }
  };

  if (menus?.length) {
    return Object.values(menus).reduce((config: MenusDefaults, menu: string | CustomMenu) => {
      const ALL_MENUS = Object.assign({}, DEFAULT);
      if (typeof menu === 'string') {
        config[menu] = ALL_MENUS[menu];
      }
      if (menu != null && typeof menu === 'object' && menu.name) {
        config[menu.name] = merge(ALL_MENUS[menu.name], menu);
      }
      return config;
    }, {});
  }
  return DEFAULT;
}


class TableMenus {
  quill: Quill;
  table: HTMLElement | null;
  root: HTMLElement;
  prevList: HTMLUListElement | null;
  prevTooltip: HTMLDivElement | null;
  scroll: boolean;
  tableBetter: QuillTableBetter;
  tablePropertiesForm: TablePropertiesForm;
  tableHeaderRow: HTMLElement | null;
  private keyboardWasVisible = false;
  private lastWindowHeight = 0;
  lastTableCreationTime = 0;
  private currentTextChangeHandler: (() => void) | null = null;
  private cellBeforeDropdown: HTMLElement | null = null;
  private isDropdownOpen = false;
  lastCellClickTime = 0;
  lastMenuShowTime = 0; // Track when menu was last shown to prevent immediate hide
  isProgrammaticScroll = false; // Flag to prevent menu hide during programmatic scroll (cell tap scroll)
  isScrollingToCell = false; // Flag to prevent double scroll - allow only one scroll mechanism
  constructor(quill: Quill, tableBetter?: QuillTableBetter) {
    this.quill = quill;
    this.table = null;
    this.prevList = null;
    this.prevTooltip = null;
    this.scroll = false;
    this.tableBetter = tableBetter;
    this.tablePropertiesForm = null;
    this.tableHeaderRow = null;
    this.quill.root.addEventListener('click', this.handleClick.bind(this));
    
    // Add keyboard event listeners to detect when keyboard is shown/hidden
    if (typeof window !== 'undefined') {
      // For iOS devices, we can detect keyboard visibility changes by listening to window resize events
      window.addEventListener('resize', this.handleWindowResize.bind(this));
    }

    // Prevent iOS contextual menu on table cells
    if (typeof window !== 'undefined' && /iPad|iPhone|iPod/.test(navigator.userAgent)) {
      // Only prevent the contextmenu event, don't interfere with touch/click
      this.quill.root.addEventListener('contextmenu', (e: Event) => {
        const target = e.target as Element;
        const cell = target.closest('td, th');
        if (cell) {
          e.preventDefault();
          e.stopPropagation();
          return false;
        }
      }, false);
      
      // Prevent text selection callout on long press
      this.quill.root.addEventListener('selectstart', (e: Event) => {
        const target = e.target as Element;
        const cell = target.closest('td, th');
        // Only prevent if cell is not focused (not being edited)
        if (cell && !cell.classList.contains('ql-cell-focused')) {
          e.preventDefault();
          return false;
        }
      }, false);
      
      // Handle touchend on table cells to show table menu instead of contextual menu
      this.quill.root.addEventListener('touchend', (e: TouchEvent) => {
        const target = e.target as Element;
        const cell = target.closest('td, th');
        if (cell) {
          // Prevent the default behavior that would show the contextual menu
          e.preventDefault();
          e.stopPropagation();
          
          // Call handleClick directly instead of dispatching synthetic event
          const mouseEvent = new MouseEvent('click', {
            bubbles: true,
            cancelable: true,
            view: window
          });
          // Set the target to the cell element
          Object.defineProperty(mouseEvent, 'target', { value: cell });
          this.handleClick(mouseEvent);
          return false;
        }
      }, false);

    }
    
    this.root = this.createMenus();
    
    // SAFE FIX: Use Quill's selection-change event to detect when cursor moves outside table
    // This is safe because it doesn't touch any mouse/touch events that could break cursor behavior
    // It only listens to Quill's own internal selection changes
    this.quill.on('selection-change', (range: any, oldRange: any, source: string) => {
      // Skip during programmatic scroll from cell tap
      if (this.isProgrammaticScroll) {
        console.log('🔍 SELECTION-CHANGE: Programmatic scroll active, skipping hide check');
        return;
      }
      
      // Skip during scrolling to cell
      if (this.isScrollingToCell) {
        console.log('🔍 SELECTION-CHANGE: Scrolling to cell, skipping hide check');
        return;
      }
      
      // If editor lost focus entirely, hide menu
      if (!range) {
        console.log('🔍 SELECTION-CHANGE: Editor lost focus, hiding menu');
        this.hideMenus();
        this.tableBetter.cellSelection.clearSelected();
        return;
      }
      
      // Check if new selection is inside a table cell
      // Wrapped in try/catch to fail silently and never break cursor behavior
      try {
        const [leaf] = this.quill.getLeaf(range.index);
        const domNode = leaf?.domNode as HTMLElement;
        const isInTable = domNode?.closest('td') !== null
                       || domNode?.closest('th') !== null
                       || domNode?.closest('table') !== null;
        
        if (!isInTable) {
          console.log('🔍 SELECTION-CHANGE: Selection moved outside table, hiding menu');
          this.hideMenus();
          this.tableBetter.cellSelection.clearSelected();
        }
      } catch (e) {
        // Fail silently - never crash cursor behavior
        console.log('🔍 SELECTION-CHANGE: Error checking table position, ignoring');
          this.hideMenus();
          this.tableBetter.cellSelection.clearSelected();
      }
    });

    // CRITICAL FIX: Listen to text-change event to hide menu when user types outside table
    // This catches typing that doesn't move the cursor (selection-change won't fire)
    // Uses Quill blot index range instead of DOM traversal (more reliable on iOS)
    this.quill.on('text-change', (delta: any, oldDelta: any, source: string) => {
      if (source !== 'user') return;
      
      // Menu already hidden — nothing to do
      if (this.root.classList.contains('ql-hidden')) return;
      
      // No table reference — safe to hide
      if (!this.table) {
        this.hideMenus(true);
        return;
      }
      
      try {
        // Get change position from delta
        let changeIndex = 0;
        if (delta.ops) {
          for (const op of delta.ops) {
            if (op.retain) changeIndex += op.retain;
            else break;
          }
        }
        
        // Use Quill's blot system to get table's start/end index
        // This is far more reliable than DOM leaf traversal on iOS
        const tableBlot = this.quill.scroll.find(this.table);
        if (!tableBlot) {
          console.log('🔍 TEXT-CHANGE: Table blot not found, hiding menu');
          this.hideMenus(true);
          this.tableBetter.cellSelection.clearSelected();
          return;
        }
        
        const tableStart = this.quill.getIndex(tableBlot);
        const tableEnd = tableStart + tableBlot.length();
        
        const isInTable = changeIndex >= tableStart && changeIndex <= tableEnd;
        
        console.log('🔍 TEXT-CHANGE: changeIndex', changeIndex, 'tableStart', tableStart, 'tableEnd', tableEnd, 'isInTable', isInTable);
        
        if (!isInTable) {
          console.log('🔍 TEXT-CHANGE: Typing outside table, hiding menu');
          this.hideMenus(true);
          this.tableBetter.cellSelection.clearSelected();
        }
      } catch (e) {
        console.log('🔍 TEXT-CHANGE: Error -', e);
        // Don't hide on error — better to leave visible than wrongly hide
      }
    });
    
    // Add document-level click listener to hide menus when clicking outside the table
    document.addEventListener('click', this.handleDocumentClick.bind(this), true);
    
    // Add touchstart listener for mobile devices (fires before click)
    // This ensures menu hides immediately on touch, not after 300ms click delay
    document.addEventListener('touchstart', this.handleDocumentClick.bind(this), { capture: true, passive: true });
    
    // Add scroll listener to hide menus when scrolling outside the table
    document.addEventListener('scroll', this.handleDocumentScroll.bind(this), true);
  }

  isPerformingTableOperation = false;

  formatsToPersist = [
    'bold', 'italic', 'underline', 'strike',
    'font', 'size',
    'color', 'background',
    'align'
  ];

  convertToRow() {
    const tableBlot = Quill.find(this.table) as TableContainer;
    const tbody = tableBlot.tbody();
    const ref = tbody.children.head;
    const rows = this.getCorrectRows();
    let row = rows[0].next;
    while (row) {
      rows.unshift(row);
      row = row.next;
    }
    for (const row of rows) {
      const tdRow = this.quill.scroll.create(TableRow.blotName) as TableRow;
      row.children.forEach(th => {
        const tdFormats = th.formats()[th.statics.blotName];
        const domNode = th.domNode.cloneNode(true);
        const td = this.quill.scroll.create(domNode).replaceWith(TableCell.blotName, tdFormats);
        tdRow.insertBefore(td, null);
      });
      tbody.insertBefore(tdRow, ref);
      row.remove();
    }
    // @ts-expect-error
    const [td] = tbody.descendant(TableCell);
    this.tableBetter.cellSelection.setSelected(td.domNode);
  }

  convertToHeaderRow() {
    const tableBlot = Quill.find(this.table) as TableContainer;
    let thead = tableBlot.thead();
    if (!thead) {
      const tbody = tableBlot.tbody();
      thead = this.quill.scroll.create(TableThead.blotName) as TableThead;
      tableBlot.insertBefore(thead, tbody);
    }
    const rows = this.getCorrectRows();
    let row = rows[0].prev;
    while (row) {
      rows.unshift(row);
      row = row.prev;
    }
    for (const row of rows) {
      const thRow = this.quill.scroll.create(TableThRow.blotName) as TableThRow;
      row.children.forEach(td => {
        const tdFormats = td.formats()[td.statics.blotName];
        const domNode = td.domNode.cloneNode(true);
        const th = this.quill.scroll.create(domNode).replaceWith(TableTh.blotName, tdFormats);
        thRow.insertBefore(th, null);
      });
      thead.insertBefore(thRow, null);
      row.remove();
    }
    // @ts-expect-error
    const [th] = thead.descendant(TableTh);
    this.tableBetter.cellSelection.setSelected(th.domNode);
  }

  async copyTable() {
    if (!this.table) return;
    const tableBlot = Quill.find(this.table) as TableContainer;
    if (!tableBlot) return;
    const html = '<p><br></p>' + tableBlot.getCopyTable();
    const text = this.tableBetter.cellSelection.getText(html);
    const clipboardItem = new ClipboardItem({
      'text/html': new Blob([html], { type: 'text/html' }),
      'text/plain': new Blob([text], { type: 'text/plain' })
    });
    try {
      await navigator.clipboard.write([clipboardItem]);
      const index = this.quill.getIndex(tableBlot);
      const length = tableBlot.length();
      this.quill.setSelection(index + length, Quill.sources.SILENT);
      this.tableBetter.hideTools();
      this.quill.scrollSelectionIntoView();
    } catch (error) {
      console.error('Failed to copy table:', error);
    }
  }

  createList(children: Children) {
    if (!children) return null;
    const container = document.createElement('ul');
    for (const [, child] of Object.entries(children)) {
      const { content, divider, createSwitch, handler } = child;
      const list = document.createElement('li');
      if (createSwitch) {
        list.classList.add('ql-table-header-row');
        list.appendChild(this.createSwitch(content));
        this.tableHeaderRow = list;
      } else {
        list.innerText = content;
      }
      list.addEventListener('click', handler.bind(this));
      container.appendChild(list);
      if (divider) {
        const dividerLine = document.createElement('li');
        dividerLine.classList.add('ql-table-divider');
        container.appendChild(dividerLine);
      }
    }
    container.classList.add('ql-table-dropdown-list', 'ql-hidden');
    return container;
  }

  createMenu(left: string, right: string, isDropDown: boolean, category: string) {
    const container = document.createElement('div');
    const dropDown = document.createElement('span');
    if (isDropDown) {
      dropDown.innerHTML = left + right;
    } else {
      dropDown.innerHTML = left;
    }
    container.classList.add('ql-table-dropdown');
    dropDown.classList.add('ql-table-tooltip-hover');
    container.setAttribute('data-category', category);
    container.appendChild(dropDown);
    return container;
  }

  createMenus() {
    const { language, options = {} } = this.tableBetter;
    const { menus } = options;
    const useLanguage = language.useLanguage.bind(language);
    const container = document.createElement('div');
    container.classList.add('ql-table-menus-container', 'ql-hidden');
    for (const [category, val] of Object.entries(getMenusConfig(useLanguage, menus))) {
      const { content, icon, children, handler } = val;
      const list = this.createList(children);
      const tooltip = createTooltip(content);
      const menu = this.createMenu(icon, downIcon, !!children, category);
      menu.appendChild(tooltip);
      list && menu.appendChild(list);
      container.appendChild(menu);
      menu.addEventListener('click', handler.bind(this, list, tooltip));
    }
    this.quill.container.appendChild(container);
    return container;
  }

  createSwitch(content: string) {
    const fragment = document.createDocumentFragment();
    const title = document.createElement('span');
    const switchContainer = document.createElement('span');
    const switchInner = document.createElement('span');
    title.innerText = content;
    switchContainer.classList.add('ql-table-switch');
    switchInner.classList.add('ql-table-switch-inner');
    switchInner.setAttribute('aria-checked', 'false');
    switchContainer.appendChild(switchInner);
    fragment.append(title, switchContainer);
    return fragment;
  }

  deleteColumn(isKeyboard: boolean = false) {
    const { computeBounds, leftTd, rightTd } = this.getSelectedTdsInfo();
    const bounds = this.table.getBoundingClientRect();
    const selectTds = getComputeSelectedTds(computeBounds, this.table, this.quill.container, 'column');
    const deleteCols = getComputeSelectedCols(computeBounds, this.table, this.quill.container);
    const tableBlot = (Quill.find(leftTd) as TableCell).table();
    const { changeTds, selTds } = this.getCorrectTds(selectTds, computeBounds, leftTd, rightTd);
    if (isKeyboard && selTds.length !== this.tableBetter.cellSelection.selectedTds.length) return;
    this.tableBetter.cellSelection.updateSelected('column');
    tableBlot.deleteColumn(changeTds, selTds, this.deleteTable.bind(this), deleteCols);
    updateTableWidth(this.table, bounds, computeBounds.left - computeBounds.right);
    this.showMenus();
  }

  deleteRow(isKeyboard: boolean = false) {
    const selectedTds = this.tableBetter.cellSelection.selectedTds;
    const rows = this.getCorrectRows();
    if (isKeyboard) {
      const sum = rows.reduce((sum: number, row: TableRow) => {
        return sum += row.children.length;
      }, 0);
      if (sum !== selectedTds.length) return;
    }
    this.tableBetter.cellSelection.updateSelected('row');
    const tableBlot = (Quill.find(selectedTds[0]) as TableCell).table();
    tableBlot.deleteRow(rows, this.deleteTable.bind(this));
    this.showMenus();
  }

  deleteTable() {
    const tableBlot = Quill.find(this.table) as TableContainer;
    if (!tableBlot) return;
    const offset = tableBlot.offset(this.quill.scroll);
    const index = this.quill.getIndex(tableBlot);
    tableBlot.remove();
    // Clean up any orphaned table-related elements
    // This ensures no leftover ql-table-block elements remain
    setTimeout(() => {
      const orphanedBlocks = this.quill.root.querySelectorAll('.ql-table-block[data-cell]');
      orphanedBlocks.forEach(block => {
        // Check if this block is not inside a table
        const parentTable = block.closest('table');
        if (!parentTable) {
          // This is an orphaned block, remove it
          block.remove();
        }
      });
      
      // Also clean up any orphaned table cells
      const orphanedCells = this.quill.root.querySelectorAll('td[data-row], th[data-row]');
      orphanedCells.forEach(cell => {
        const parentTable = cell.closest('table');
        if (!parentTable) {
          cell.remove();
        }
      });
      
      // Force Quill to update and clean up its internal state
      this.quill.update(Quill.sources.USER);
    }, 0);
    this.tableBetter.hideTools();
    this.quill.setSelection(offset - 1, 0, Quill.sources.USER);

    // Editor is empty, set default font and size
    // Use stored defaults or fall back to reasonable value
    // Apply default font and size formatting
    const tableModule = this.quill.getModule('table') as Table;
        
    // Get default formats from the table module
    const defaultFont = tableModule?.defaultFormats?.font || tableModule?.lastFormat?.font || 'Helvetica';
    const defaultSize = tableModule?.defaultFormats?.size || tableModule?.lastFormat?.size || '18pt';

 
    // Apply default formats at current cursor position
    // Use silent source for the first formatting to avoid multiple updates
    this.quill.format('font', defaultFont, Quill.sources.SILENT);
    this.quill.format('size', defaultSize, Quill.sources.USER); // USER source for the last one to trigger an update
 
    // Apply to a small range to ensure formats stick
    const range = this.quill.getSelection() || { index: index, length: 1 };
    this.quill.formatText(range.index, Math.max(1, range.length), {
      font: defaultFont,
      size: defaultSize
    }, Quill.sources.USER);

    const fontPickerLabel = document.querySelector('.ql-font .ql-picker-label');
    const sizePickerLabel = document.querySelector('.ql-size .ql-picker-label');
    if (fontPickerLabel) {
      fontPickerLabel.classList.add("ql-active");
      fontPickerLabel.setAttribute('data-value', `${defaultFont}`);
    }
    if (sizePickerLabel) {
      sizePickerLabel.classList.add("ql-active");
      sizePickerLabel.setAttribute('data-value', `${defaultSize}`);
    }

    // With this:
    this.isDropdownOpen = false; // ← force reset dropdown state first
    this.hideMenus(true);        // ← then hide with full style cleanup
    
    // Reset scroll flags
    this.isScrollingToCell = false;
    this.isProgrammaticScroll = false;
  }


  destroyTablePropertiesForm() {
    if (!this.tablePropertiesForm) return;
    this.tablePropertiesForm.removePropertiesForm();
    this.tablePropertiesForm = null;
  }

  disableMenu(category: string, disabled?: boolean) {
    if (!this.root) return;
    const menu = this.root.querySelector(`[data-category=${category}]`);
    if (!menu) return;
    if (disabled) {
      menu.classList.add('ql-table-disabled');
    } else {
      menu.classList.remove('ql-table-disabled');
    }
  }

  getCellsOffset(
    computeBounds: CorrectBound,
    bounds: CorrectBound,
    leftColspan: number,
    rightColspan: number
  ) {
    const tableBlot = Quill.find(this.table) as TableContainer;
    const cells = tableBlot.descendants(TableCell);
    const _left = Math.max(bounds.left, computeBounds.left);
    const _right = Math.min(bounds.right, computeBounds.right);
    const map: TableCellMap = new Map();
    const leftMap: TableCellMap = new Map();
    const rightMap: TableCellMap = new Map();
    for (const cell of cells) {
      const { left, right } = getCorrectBounds(cell.domNode, this.quill.container);
      if (left + DEVIATION >= _left && right <= _right + DEVIATION) {
        this.setCellsMap(cell, map);
      } else if (
        left + DEVIATION >= computeBounds.left &&
        right <= bounds.left + DEVIATION
      ) {
        this.setCellsMap(cell, leftMap);
      } else if (
        left + DEVIATION >= bounds.right &&
        right <= computeBounds.right + DEVIATION
      ) {
        this.setCellsMap(cell, rightMap);
      }
    }
    return this.getDiffOffset(map) ||
      this.getDiffOffset(leftMap, leftColspan)
      + this.getDiffOffset(rightMap, rightColspan);
  }

  getColsOffset(
    colgroup: TableColgroup,
    computeBounds: CorrectBound,
    bounds: CorrectBound
  ) {
    let col = colgroup.children.head;
    const _left = Math.max(bounds.left, computeBounds.left);
    const _right = Math.min(bounds.right, computeBounds.right);
    let colLeft = null;
    let colRight = null;
    let offset = 0;
    while (col) {
      const { width } = col.domNode.getBoundingClientRect();
      if (!colLeft && !colRight) {
        const colBounds = getCorrectBounds(col.domNode, this.quill.container);
        colLeft = colBounds.left;
        colRight = colLeft + width;
      } else {
        colLeft = colRight;
        colRight += width;
      }
      if (colLeft > _right) break;
      if (colLeft >= _left && colRight <= _right) {
        offset--;
      }
      col = col.next;
    }
    return offset;
  }

  getCorrectBounds(table: HTMLElement): CorrectBound[] {
    const bounds = this.quill.container.getBoundingClientRect();
    const tableBounds = getCorrectBounds(table, this.quill.container);
    return (
      tableBounds.width >= bounds.width
       ? [{ ...tableBounds, left: 0, right: bounds.width }, bounds]
       : [tableBounds, bounds]
    );
  }

  getCorrectTds(
    selectTds: Element[],
    computeBounds: CorrectBound,
    leftTd: Element,
    rightTd: Element
  ) {
    const changeTds: [Element, number][] = [];
    const selTds = [];
    const colgroup = (Quill.find(leftTd) as TableCell).table().colgroup() as TableColgroup;
    const leftColspan = (~~leftTd.getAttribute('colspan') || 1);
    const rightColspan = (~~rightTd.getAttribute('colspan') || 1);
    if (colgroup) {
      for (const td of selectTds) {
        const bounds = getCorrectBounds(td, this.quill.container);
        if (
          bounds.left + DEVIATION >= computeBounds.left &&
          bounds.right <= computeBounds.right + DEVIATION
        ) {
          selTds.push(td);
        } else {
          const offset = this.getColsOffset(colgroup, computeBounds, bounds);
          changeTds.push([td, offset]);
        }
      }
    } else {
      for (const td of selectTds) {
        const bounds = getCorrectBounds(td, this.quill.container);
        if (
          bounds.left + DEVIATION >= computeBounds.left &&
          bounds.right <= computeBounds.right + DEVIATION
        ) {
          selTds.push(td);
        } else {
          const offset = this.getCellsOffset(
            computeBounds,
            bounds,
            leftColspan,
            rightColspan
          );
          changeTds.push([td, offset]);
        }
      }
    }
    return { changeTds, selTds };
  }

  getCorrectRows() {
    const selectedTds = this.tableBetter.cellSelection.selectedTds;
    const map: { [propName: string]: TableRow } = {};
    for (const td of selectedTds) {
      let rowspan = ~~td.getAttribute('rowspan') || 1;
      let row = Quill.find(td.parentElement) as TableRow;
      if (rowspan > 1) {
        while (row && rowspan) {
          const id = row.children.head.domNode.getAttribute('data-row');
          if (!map[id]) map[id] = row;
          row = row.next;
          rowspan--;
        }
      } else {
        const id = td.getAttribute('data-row');
        if (!map[id]) map[id] = row;
      }
    }
    return Object.values(map);
  }

  getDiffOffset(map: TableCellMap, colspan?: number) {
    let offset = 0;
    const tds = this.getTdsFromMap(map);
    if (tds.length) {
      if (colspan) {
        for (const td of tds) {
          offset += (~~td.getAttribute('colspan') || 1);
        }
        offset -= colspan;
      } else {
        for (const td of tds) {
          offset -= (~~td.getAttribute('colspan') || 1);
        }
      }
    }
    return offset;
  }

  getRefInfo(row: TableRow, right: number) {
    let ref = null;
    if (!row) return { id: tableId(), ref };
    let td = row.children.head;
    const id = td.domNode.getAttribute('data-row');
    while (td) {
      const { left } = td.domNode.getBoundingClientRect();
      if (Math.abs(left - right) <= DEVIATION) {
        return { id, ref: td };
        // The nearest cell of a multi-row cell
      } else if (Math.abs(left - right) >= DEVIATION && !ref) {
        ref = td;
      }
      td = td.next;
    }
    return { id, ref };
  }

  getSelectedTdAttrs(td: HTMLElement) {
    const cellBlot = Quill.find(td) as TableCell;
    const align = getAlign(cellBlot);
    const attr: Props =
      align
        ? { ...getElementStyle(td, CELL_PROPERTIES), 'text-align': align }
        : getElementStyle(td, CELL_PROPERTIES);
    return attr;
  }

  getSelectedTdsAttrs(selectedTds: HTMLElement[]) {
    const map = new Map();
    let attribute = null;
    for (const td of selectedTds) {
      const attr = this.getSelectedTdAttrs(td);
      if (!attribute) {
        attribute = attr;
        continue;
      }
      for (const key of Object.keys(attribute)) {
        if (map.has(key)) continue;
        if (attr[key] !== attribute[key]) {
          map.set(key, false);
        }
      }
    }
    for (const key of Object.keys(attribute)) {
      if (map.has(key)) {
        attribute[key] = CELL_DEFAULT_VALUES[key];
      }
    }
    return attribute;
  }

  getSelectedTdsInfo() {
    // PRIMARY: Use lastClickedCell — most reliable, set directly on tap
    let startTd = this.tableBetter.lastClickedCell as HTMLElement;
    let endTd = this.tableBetter.cellSelection?.endTd as HTMLElement;

    // Validate startTd
    if (!startTd || !startTd.isConnected) {
      console.log('lastClickedCell invalid, falling back to cellSelection.startTd');
      startTd = this.tableBetter.cellSelection?.startTd as HTMLElement;
      
      if (!startTd || !startTd.isConnected) {
        console.log('cellSelection.startTd invalid, attempting reconnect...');
        startTd = this.reconnectCellReference(startTd as HTMLElement);
      }

      if (!startTd || !startTd.isConnected) {
        console.log('Reconnect failed, trying selectedTds...');
        // Last resort: selectedTds
        const selected = this.tableBetter.cellSelection?.selectedTds;
        if (selected?.length > 0 && selected[0].isConnected) {
          startTd = selected[0] as HTMLElement;
          console.log('Using first selectedTd');
        }
      }
    }

    // Validate endTd
    if (!endTd || !endTd.isConnected) {
      console.log('endTd invalid, attempting reconnect...');
      endTd = this.reconnectCellReference(endTd as HTMLElement);
      if (!endTd || !endTd.isConnected) {
        console.log('endTd reconnect failed, using startTd as fallback');
        endTd = startTd; // fallback to startTd
      }
    }

    if (!startTd || !startTd.isConnected) {
      console.error('Unable to find valid cells for table operation');
      return null;
    }

    const startCorrectBounds = getCorrectBounds(startTd, this.quill.container);
    const endCorrectBounds = getCorrectBounds(endTd, this.quill.container);
    const computeBounds = getComputeBounds(startCorrectBounds, endCorrectBounds);

    // Determine geometric left/right and top/bottom cells separately
    const isStartTop  = startCorrectBounds.top  <= endCorrectBounds.top;
    const isStartLeft = startCorrectBounds.left <= endCorrectBounds.left;

    return {
      computeBounds,
      leftTd:   isStartLeft ? startTd : endTd,
      rightTd:  isStartLeft ? endTd   : startTd,
      topTd:    isStartTop  ? startTd : endTd,
      bottomTd: isStartTop  ? endTd   : startTd,
    };
  }

  getTableAlignment(table: HTMLTableElement) {
    const align = table.getAttribute('align');
    if (!align) {
      const {
        [Alignment.left]: left,
        [Alignment.right]: right
      } = getElementStyle(table, [Alignment.left, Alignment.right]);
      if (left === 'auto') {
        if (right === 'auto') return 'center';
        return 'right';
      }
      return 'left';
    }
    return align || 'center';
  }

  /**
   * Refreshes the cell selection after table operations to ensure we have valid DOM references.
   * This is called after insert/delete operations that may have invalidated previous references.
   */
  refreshCellSelection() {
    try {
      const { cellSelection } = this.tableBetter;
      const { selectedTds } = cellSelection;

      // Check if any selected cells are disconnected
      const validCells = selectedTds.filter(td => td && td.isConnected);
      const hasDisconnectedCells = validCells.length !== selectedTds.length;

      if (hasDisconnectedCells) {
        console.log('Found disconnected cells in selection, refreshing...');

        if (validCells.length > 0) {
          // We still have some valid cells, update the selection
          cellSelection.setSelectedTds(validCells);
        } else {
          // All cells are disconnected, try to find a valid cell in the table
          const table = this.table;
          if (table) {
            const firstCell = table.querySelector('td, th') as HTMLElement;
            if (firstCell && firstCell.isConnected) {
              console.log('No valid cells in selection, selecting first available cell');
              cellSelection.setSelectedTds([firstCell]);
            } else {
              console.warn('Unable to find any valid cells in table after refresh');
            }
          }
        }
      }

      // Update the internal startTd and endTd references
      const { selectedTds: updatedSelectedTds } = cellSelection;
      if (updatedSelectedTds.length > 0) {
        cellSelection.startTd = updatedSelectedTds[0];
        cellSelection.endTd = updatedSelectedTds[updatedSelectedTds.length - 1] || updatedSelectedTds[0];
      }
    } catch (error) {
      console.error('Error refreshing cell selection:', error);
    }
  }

  getTdsFromMap(map: TableCellMap) {
    return Object.values(Object.fromEntries(map))
    .reduce((tds: HTMLTableCellElement[], item: HTMLTableCellElement[]) => {
      return tds.length > item.length ? tds : item;
    }, []);
  }

  handleIOSTableOperation(type: string, offset: number) {
    // Check if we have a valid last clicked cell
    const lastClickedCell = this.tableBetter.lastClickedCell;
    if (!lastClickedCell || !lastClickedCell.isConnected) {
      return false;
    }
    
    // Check if the keyboard is visible
    if (!this.keyboardWasVisible) {
      return false;
    }
    
    // Perform the table operation
    if (type === 'column') {
      this.insertColumn(lastClickedCell, offset);
    } else if (type === 'row') {
      this.insertRow(lastClickedCell, offset);
    }
    
    // Update the menu position
    this.updateMenus();
    
    // Restore the cell selection
    this.tableBetter.cellSelection.setSelectedTds([lastClickedCell]);
    
    return true;
  }

  handleWindowResize() {
    const currentHeight = window.innerHeight;
    
    // If height decreased significantly, keyboard is likely shown
    if (this.lastWindowHeight > 0 && this.lastWindowHeight - currentHeight > 150) {
      this.keyboardWasVisible = true;
    }
    
    // If height increased significantly after keyboard was visible, keyboard is likely hidden
    if (this.keyboardWasVisible && currentHeight - this.lastWindowHeight > 150) {
      this.keyboardWasVisible = false;
      
      // Ensure menu container is visible after keyboard is hidden
      if (this.table) {
        // Force a delay to ensure proper rendering after keyboard hide
        setTimeout(() => {
          try {
            // Get the last clicked cell from the tableBetter module
            const lastClickedCell = this.tableBetter.lastClickedCell;
            
            // If we have a valid last clicked cell, select it
            if (lastClickedCell && lastClickedCell.isConnected) {
              // First clear any existing selection
              this.tableBetter.cellSelection.clearSelected();
              
              // Then programmatically select the cell
              this.tableBetter.cellSelection.setSelectedTds([lastClickedCell]);
              
              console.log('Re-selected cell after keyboard hide:', lastClickedCell);
              
              // Show menus and update their position
              this.showMenus();
              this.updateMenus(this.table);
              
              // Make sure dropdowns are visible if they were open before
              setTimeout(() => {
                const dropdowns = document.querySelectorAll('.ql-table-dropdown');
                dropdowns.forEach(dropdown => {
                  if (dropdown.classList.contains('ql-hidden')) {
                    dropdown.classList.remove('ql-hidden');
                  }
                });
              }, 50);
            } else {
              // Just show menus without cell selection
              this.showMenus();
              this.updateMenus(this.table);
            }
          } catch (error) {
            console.error('Error restoring cell selection after keyboard hide:', error);
            // Fallback to just showing menus
            this.showMenus();
            this.updateMenus(this.table);
          }
        }, 300);
      }
    }
    
    this.lastWindowHeight = currentHeight;
  }

  handleClick(e: MouseEvent) {
    console.log("e.detail",e.detail);
     // Add detection for double-tap/double-click
    if (e.detail === 2) {
      // For double-taps, only allow text selection
      // but prevent any table structure changes
      return;
    }
    if (!this.quill.isEnabled()) return;
    let table = (e.target as Element).closest('table');
    if (table && !this.quill.root.contains(table)) return;
    
    // Store the clicked element to check if it's a menu icon
    const clickedElement = e.target as Element;
    const isMenuIconClick = clickedElement.closest('.ql-table-dropdown') !== null;
    
    // Save current cell selection before any UI changes
    const currentSelectedTds = [...this.tableBetter.cellSelection.selectedTds];
    const hasSelection = currentSelectedTds.length > 0;
    
    this.prevList && this.prevList.classList.add('ql-hidden');
    this.prevTooltip && this.prevTooltip.classList.remove('ql-table-tooltip-hidden');
    this.prevList = null;
    this.prevTooltip = null;
    
    // Check if we're on iOS and the keyboard might be visible
    const isIOS = typeof window !== 'undefined' && /iPad|iPhone|iPod/.test(navigator.userAgent);
    const windowHeight = window.innerHeight;
    const keyboardLikelyVisible = isIOS && windowHeight < window.outerHeight * 0.8;
    
    // Get the clicked cell
    const cell = clickedElement.closest('td, th');
    
    // Check if we're clicking on the menus container itself (dropdown, tooltip, etc.)
    const isClickingOnMenus = clickedElement.closest('.ql-table-menus-container') !== null;
    
    // If no table and no cell was clicked, hide menus
    // BUT: Don't hide if we're clicking on the menus themselves (dropdown, etc.)
    if (!table && !cell && !isClickingOnMenus) {
      this.hideMenus();
      this.destroyTablePropertiesForm();
       // Clear any table selection
      this.tableBetter.cellSelection.clearSelected();
      
      // Remove any existing table-block classes from non-table elements
      const blockElements = this.quill.root.querySelectorAll('.ql-table-block');
      blockElements.forEach(block => {
        if (!block.closest('table')) {
          block.classList.remove('ql-table-block');
        }
      });
      
      // Ensure clicking below table creates plain paragraph
      if (e.target === this.quill.root) {
        const range = this.quill.getSelection(true);
        if (range) {
          this.quill.formatLine(range.index, range.length, 'table-block', false);
        }
      }

      // FOCUS FOR AUTO-SCROLL: Use adjustScrollForSelection for scrolling
      console.log('🔍 HANDLECLICK: Calling adjustScrollForSelection for regular text click');
      if (typeof window !== 'undefined' && (window as any).adjustScrollForSelection) {
        const editor = document.getElementById('editor');
        const currentSelection = this.quill.getSelection();
        if (editor && currentSelection) {
          (window as any).adjustScrollForSelection(currentSelection.index, editor, this.quill);
        }
      }
       else {
        // Fallback: just focus the editor
        requestAnimationFrame(() => {
          this.quill.focus();
        });
      }

      return;
    }
    
    // If a cell was clicked but no table found, still don't hide menus
    // (user might be clicking on a cell within a table)
    if (cell && !table) {
      table = cell.closest('table');
      if (!table) {
        // Cell is not in a table, hide menus
        this.hideMenus();
        this.destroyTablePropertiesForm();
        return;
      }
    }
    
    // Don't process if table properties form is open
    if (this.tablePropertiesForm) return;
    
    // If a cell was clicked, handle cell selection and menu display
    if (cell && !isMenuIconClick) {
      console.log(`🔍 HANDLECLICK: Cell clicked, starting selection process`);

      // Call handlePreciseTextPositioning BEFORE stopping propagation
      // This ensures precise cursor positioning happens before any event blocking
      if (typeof window !== 'undefined' && (window as any).handlePreciseTextPositioning) {
        console.log('🔍 HANDLECLICK: Calling handlePreciseTextPositioning for precise positioning');
        (window as any).handlePreciseTextPositioning(e);
      }

      // Prevent event bubbling to toolbar buttons
      e.stopPropagation();

      // Ensure we have the actual td/th element, not a child element
      const tdElement = cell.tagName === 'TD' || cell.tagName === 'TH' 
        ? cell 
        : cell.closest('td, th');
      
      if (tdElement) {
        // Update the last clicked cell reference
        this.tableBetter.lastClickedCell = tdElement as HTMLElement;
        
        // Track when cell was clicked for scroll prevention
        this.lastCellClickTime = Date.now();
        
        // CRITICAL FIX: Ensure editor is focused IMMEDIATELY and SYNCHRONOUSLY
        // This is essential for first-tap guarantee after keyboard dismissal
        // We do this BEFORE setSelectedTds to ensure Quill is ready
        // Check multiple conditions to handle all blur states
        const isEditorFocused = document.activeElement === this.quill.root;
        const isEditorEnabled = this.quill.isEnabled && this.quill.isEnabled();
        
        if (!isEditorFocused || !isEditorEnabled) {
          console.log('🔍 HANDLECLICK: Editor not focused/enabled, restoring state');
          console.log('  - isEditorFocused:', isEditorFocused);
          console.log('  - isEditorEnabled:', isEditorEnabled);
          
          // Enable editor if disabled (happens after keyboard "Done" button)
          if (!isEditorEnabled) {
            this.quill.enable(true);
            console.log('🔍 HANDLECLICK: Editor enabled');
          }
          
          // Focus editor synchronously (not via requestAnimationFrame)
          this.quill.focus();
          console.log('🔍 HANDLECLICK: Editor focused synchronously');
        }
        
        // STEP 2: Set the cell selection FIRST (this will apply ql-cell-focused class immediately)
        // This ensures the blue highlight appears on first tap
        this.tableBetter.cellSelection.setSelectedTds([tdElement as HTMLElement]);
        
        // NOTE: Do NOT show menu here - it will be shown ONLY once after scroll completes
        // This prevents the menu from flashing at the wrong position before scroll
        // The menu will be shown in updateMenus() after scrollend event fires with fresh coordinates
        
        // STEP 4: Do NOT clear deselection flags when clicking a new cell
        // Deselection is a GLOBAL user preference that persists across all cells
        // When user deselects bold, it should stay deselected everywhere until they reselect it
        
        // STEP 5: Get the calculated cursor position from handlePreciseTextPositioning if available
        const calculatedPosition = (window as any).lastCalculatedCursorPosition;
        console.log('🔍 TABLE-MENUS: Calculated position from handlePreciseTextPositioning:', calculatedPosition);
        
        // Clear the stored position after using it
        (window as any).lastCalculatedCursorPosition = undefined;
        
        // NOTE: Focus was already called synchronously before setSelectedTds()
        // No need to call it again here - that would be redundant and could cause race conditions
        
        // STEP 5: Update toolbar with actual cell formats
        setTimeout(() => {
          const range = this.quill.getSelection();
          console.log("range", range);
          if (!range || range.index === null || range.index === undefined) return;
          
          // FIRST: Check if cell has actual content (using innerText which is most reliable)
          // Important: A cell with just a newline (↵) should be considered empty
          const cellText = (tdElement as HTMLElement).innerText || '';
          // Remove zero-width spaces, trim whitespace, and check if anything meaningful remains
          const meaningfulContent = cellText
            .replace(/\u200B/g, '') // Remove zero-width spaces
            .replace(/\n/g, '') // Remove newlines
            .replace(/\r/g, '') // Remove carriage returns
            .trim();
          const hasContent = meaningfulContent.length > 0;
          console.log("🔍 TABLE-MENUS: Cell text check - hasContent:", hasContent, "cellText:", JSON.stringify(cellText), "meaningful:", JSON.stringify(meaningfulContent));
          
          // Get the ACTUALLY CLICKED CELL (tdElement), not from getTable()[2]
          const clickedCellBlot = Quill.find(tdElement) as any;
          let formats: any = {};
          
          if (clickedCellBlot) {
            // Get formats from the clicked cell
            const cellFormats = this.tableBetter.getCellTextFormats(clickedCellBlot);
            // REQUIREMENT: Remove formats from userDeselectedFormats if they're already active in this cell
            // If a format is applied to the cell content, it shouldn't be in the deselected list
            this.tableBetter.removeDeselectedFormatsIfActive(cellFormats);
            
            if (hasContent) {
              // NON-EMPTY CELL: Use actual cell formats
              formats = { ...cellFormats};
            } else {
              // EMPTY CELL: Priority order for format inheritance
              // 1. FormatManager.globalFormats (HIGHEST - what user selected in toolbar)
              // 2. Previous cell formats (FALLBACK - for formats not in globalFormats)
              // Get global formats from FormatManager (what user selected in toolbar)
              const globalFormats = this.tableBetter.formatManager.getGlobalFormats();
              
              // Start with global formats as base (highest priority)
              formats = { ...globalFormats };
              
              // Get previous cell element
              const table = tdElement.closest('table');
              const cells = table ? Array.from(table.querySelectorAll('td, th')) : [];
              const currentIndex = cells.indexOf(tdElement);
              const prevCell = currentIndex > 0 ? cells[currentIndex - 1] : null;
              
              if (prevCell) {
                // Read font directly from DOM skipping empty spans
                const rockwellSpan = prevCell.querySelector('[class*="ql-font-"]:not(:empty)') 
                  || Array.from(prevCell.querySelectorAll('[class*="ql-font-"]'))
                     .find((el: Element) => {
                       const text = el.textContent?.replace(/\u200B/g, '').trim() || '';
                       return text.length > 0;
                     });
                
                if (rockwellSpan) {
                  const fontClass = Array.from(rockwellSpan.classList)
                    .find((c: string) => c.startsWith('ql-font-'));
                  if (fontClass) {
                    formats.font = fontClass.replace('ql-font-', '');
                    console.log("🔍 TABLE-MENUS: Read font from previous cell DOM:", formats.font);
                  }
                }
                
                // Read size from inline style or class
                const styledEl = prevCell.querySelector('[style*="font-size"]') ||
                  prevCell.querySelector('[class*="ql-size-"]:not(:empty)') ||
                  Array.from(prevCell.querySelectorAll('[class*="ql-size-"]'))
                    .find((el: Element) => {
                      const text = el.textContent?.replace(/\u200B/g, '').trim() || '';
                      return text.length > 0;
                    });
                
                if (styledEl) {
                  const htmlEl = styledEl as HTMLElement;
                  if (htmlEl.style?.fontSize) {
                    formats.size = htmlEl.style.fontSize;
                  } else {
                    const sizeClass = Array.from(styledEl.classList)
                      .find((c: string) => c.startsWith('ql-size-'));
                    if (sizeClass) {
                      formats.size = sizeClass.replace('ql-size-', '');
                    }
                  }
                  console.log("🔍 TABLE-MENUS: Read size from previous cell DOM:", formats.size);
                }
              }
              
              // CRITICAL: Global formats ALWAYS have highest priority - override everything
              // This ensures FormatManager.globalFormats is respected for ALL formats including font/size
              formats = { ...formats, ...globalFormats };
              console.log("🔍 TABLE-MENUS: Final formats after global override:", formats);
              
              // Try lastAppliedFormats as fallback for inline formats only
              const lastApplied = this.tableBetter.getLastAppliedFormats();
              
              if (Object.keys(lastApplied).length > 0) {
                // Only merge inline formats from lastApplied (font/size already handled by DOM reading)
                const mergedFormats = { ...formats };
                
                const inlineFormats = ['bold', 'italic', 'underline', 'strike'];
                inlineFormats.forEach(format => {
                  if (lastApplied[format] !== undefined && !globalFormats[format]) {
                    mergedFormats[format] = lastApplied[format];
                    console.log(`🔍 TABLE-MENUS: Inherited inline format ${format} from lastApplied:`, lastApplied[format]);
                  }
                });
                
                formats = mergedFormats;
              }
              
              // IMPORTANT: Only remove inline deselected formats from empty cell
              // Font and size are structural formats that should ALWAYS be preserved
              const deselectedFormats = this.tableBetter.formatManager.getDeselectedFormats();
              const deselectedInlineFormats = ['bold', 'italic', 'underline', 'strike'];
              Object.keys(deselectedFormats).forEach(format => {
                if (deselectedFormats[format] && deselectedInlineFormats.includes(format)) {
                  delete formats[format];
                }
              });
            }
          }
          
          // Filter out table metadata
          delete formats['table-cell-block'];
          delete formats['table-cell'];

          this.tableBetter.updateToolbarUI(formats, hasContent);
          
          // Update toolbar using consolidated sync function
          if ((window as any).syncCellToolbar) {
            (window as any).syncCellToolbar(formats);
          }
        }, 150);

        // STEP 6: Scroll cell into view using adjustScrollForSelection
        // Reuses proven scroll logic from HTML file instead of duplicating
        setTimeout(() => {
          console.log('🔍 TABLE-MENUS: Scrolling cell to center');

          if (this.isScrollingToCell) {
            console.log('🔍 TABLE-MENUS: Scroll already in progress, skipping');
            return;
          }

          this.isScrollingToCell = true;
          this.isProgrammaticScroll = true;

          // Get the editor container (parent of quill.root)
          const editorContainer = this.quill.root.parentElement || document.getElementById('editor');
          
          // Call adjustScrollForSelection if available
          const selection = this.quill.getSelection();
          if (selection && editorContainer && (window as any).adjustScrollForSelection) {
            console.log('🔍 TABLE-MENUS: Calling adjustScrollForSelection with index', selection.index);
            (window as any).adjustScrollForSelection(selection.index, editorContainer, this.quill);
          }

          // Reposition menu after scroll settles
          setTimeout(() => {
            console.log('🔍 TABLE-MENUS: Scroll complete, repositioning menu');
            if (table && tdElement.isConnected) {
              this.updateMenus(table, tdElement as HTMLElement);
            }
            this.isScrollingToCell = false;
            this.isProgrammaticScroll = false;
            console.log('🔍 TABLE-MENUS: Set isScrollingToCell = false, isProgrammaticScroll = false');
          }, 250); // enough time for adjustScrollForSelection's internal 100ms + scroll to finish

        }, 100);
        
      }
    } else if (isMenuIconClick && hasSelection) {
      // When clicking on a dropdown icon, preserve the current cell selection
      // Don't change the selection, just keep the current one active
      if (currentSelectedTds.length > 0 && currentSelectedTds[0].isConnected) {
        // Ensure the cell remains focused
        currentSelectedTds[0].classList.add('ql-cell-focused');
      }
    }
    
    // Show menus if we have a table or selected cells (menu already shown above, just update position)
    if (table || this.tableBetter.cellSelection.selectedTds.length > 0) {
      // Menu was already shown before setSelectedTds(), just update position now
      
      // Update table reference if we have one
      if (table) {
        if (!table.isEqualNode(this.table) || this.scroll) {
          this.updateScroll(false);
        }
        this.table = table;
        // Always pass the clicked cell (tdElement) to updateMenus for accurate positioning
        // This ensures the menu appears relative to the cell that was actually clicked
        const cellToPosition = (cell && cell.tagName === 'TD' || cell?.tagName === 'TH') ? cell as HTMLElement : null;
        this.updateMenus(table, cellToPosition || this.tableBetter.cellSelection.selectedTds[0] as HTMLElement);
      } else if (this.tableBetter.cellSelection.selectedTds.length > 0) {
        // Try to find table from selected cells
        const selectedCell = this.tableBetter.cellSelection.selectedTds[0];
        const tableFromCell = selectedCell?.closest('table');
        if (tableFromCell) {
          this.table = tableFromCell;
          // Pass the selected cell to updateMenus for accurate positioning
          this.updateMenus(tableFromCell, selectedCell as HTMLElement);
        }
      }
    }
    
    // Handle iOS keyboard visibility restoration - immediate approach for older iOS
    // if (keyboardLikelyVisible && hasSelection && currentSelectedTds[0]?.isConnected) {
    //   // Use version-aware keyboard restoration
    //   if (typeof window !== 'undefined' && (window as any).shouldRestoreKeyboardOnCellSwitch) {
    //     const shouldRestore = (window as any).shouldRestoreKeyboardOnCellSwitch();
    //     if (shouldRestore) {
    //       console.log('🔍 Older iOS detected - immediately restoring keyboard focus');

    //       // Immediately restore focus for older iOS - no delay needed
    //       // On older iOS, the keyboard doesn't automatically reopen when focus changes
    //       requestAnimationFrame(() => {
    //         console.log('🔍 Executing keyboard focus restoration');
    //         this.quill.focus();
    //       });
    //     }
    //   }
    // }
  }

  handleDocumentClick(e: Event) {
    const target = e.target as Element;
    
    // Don't hide if clicking on the menus themselves
    if (this.root.contains(target)) {
      console.log('🔍 HANDLEDOCUMENTCLICK: Click is on menu, not hiding');
      return;
    }
    
    // Don't hide if clicking on the table properties form
    if (this.tablePropertiesForm && this.tablePropertiesForm.form && this.tablePropertiesForm.form.contains(target)) {
      console.log('🔍 HANDLEDOCUMENTCLICK: Click is on table properties form, not hiding');
      return;
    }
    
    // If clicking ANYWHERE inside ANY table — let handleClick manage it entirely
    // Don't use this.table here because it may not be set yet (race condition on iOS touchstart)
    const clickedTable = target.closest('table');
    if (clickedTable) {
      console.log('🔍 HANDLEDOCUMENTCLICK: Click is inside a table, letting handleClick manage it');
      return; // ← Just bail out completely, handleClick will show/hide correctly
    }
    
    // Click is truly outside any table — safe to hide
    console.log('🔍 HANDLEDOCUMENTCLICK: Click is outside table, hiding menus');
    this.hideMenus(true); // ← true = also clear inline styles
    this.tableBetter.cellSelection.clearSelected();
  }

  handleDocumentScroll(e: Event) {
    // CRITICAL FIX: Skip ALL hide logic during programmatic scroll
    // The cell tap triggers scrollIntoView which fires scroll events during animation
    // We must not hide the menu or check visibility while scroll is in progress
    if (this.isProgrammaticScroll) {
      console.log('🔍 HANDLEDOCUMENTSCROLL: Programmatic scroll detected, skipping all hide logic');
      return;
    }
    
    // Only check "table out of view" for manual user scroll (not programmatic)
    const target = e.target as Element;
    if (this.quill.container.contains(target) || target === this.quill.container) {
      if (this.table && this.table.isConnected) {
        const tableBounds = this.table.getBoundingClientRect();
        const containerBounds = this.quill.container.getBoundingClientRect();
        
        // If table is completely out of view, hide menus
        if (tableBounds.bottom < containerBounds.top || tableBounds.top > containerBounds.bottom) {
          console.log('🔍 HANDLEDOCUMENTSCROLL: Table out of view, hiding menus');
          this.hideMenus();
        }
      }
    }
  }

  hideMenus(clearInlineStyles: boolean = false) {
    // Don't hide if a dropdown is currently open
    if (this.isDropdownOpen) {
      return;
    }
    // Just add the class — !important in CSS guarantees it beats inline styles
    this.root.classList.add('ql-hidden');
    // Only clear inline styles when user clicked outside table
    // Don't clear when called during updateMenus() repositioning flow
    if (clearInlineStyles) {
      this.root.style.visibility = '';
      this.root.style.opacity = '';
      this.root.style.display = '';
    }
  }

  getDropdownOpen(): boolean {
    return this.isDropdownOpen;
  }

   markTableAsNewlyCreated() {
    this.lastTableCreationTime = Date.now();
  }

  insertColumn(td: HTMLElement, offset: number) {
    try {
      // Validate td parameter
      if (!td) {
        console.warn('Cannot insert column: td parameter is null or undefined');
        return;
      }

      // Check if td is still connected to the DOM
      if (!td.isConnected) {
        // Try to reconnect the cell reference
        const reconnectedTd = this.reconnectCellReference(td);
        if (reconnectedTd) {
          console.log('Reconnected td reference for column insertion');
          td = reconnectedTd;
        } else {
          console.error('Cannot insert column: unable to reconnect td reference');
          return;
        }
      }
      
      // Safely get bounding client rect
      let left = 0, right = 0, width = 0;
      try {
        const rect = td.getBoundingClientRect();
        left = rect.left;
        right = rect.right;
        width = rect.width;
      } catch (error) {
        console.error('Error getting bounding client rect:', error);
        return;
      }
      
      // Find the tdBlot
      const tdBlot = Quill.find(td) as TableCell;
      if (!tdBlot) {
        console.warn('Cannot insert column: tdBlot not found');
        return;
      }
      
      // Get the table blot
      const tableBlot = tdBlot.table();
      if (!tableBlot) {
        console.warn('Cannot insert column: tableBlot not found');
        return;
      }
      
      // Check if parentElement exists
      if (!td.parentElement) {
        console.warn('Cannot insert column: td.parentElement is null');
        return;
      }
      
      // Check if lastChild exists
      if (!td.parentElement.lastChild) {
        console.warn('Cannot insert column: td.parentElement.lastChild is null');
        return;
      }
      
      const isLast = td.parentElement.lastChild.isEqualNode(td);
      const position = offset > 0 ? right : left;
      
      // Track existing cells before insertion to identify newly created ones
      const existingCells = new Set(Array.from(this.table?.querySelectorAll('td, th') || []));
      
      // Insert the column
      tableBlot.insertColumn(position, isLast, width, offset);
      
      // Apply formats to newly inserted cells
      // if (cellFormats && cellFormats[0]) {
        const allCellsAfterInsertion = Array.from(this.table?.querySelectorAll('td, th') || []);
        const newCells = allCellsAfterInsertion.filter(cell => !existingCells.has(cell));
        
        newCells.forEach((cell: any) => {
          const cellBlot: any = Quill.find(cell) as TableCell;
          if (cellBlot) {
            // Use applyFormatOnEmptyCell instead of direct format application
            if (typeof window !== 'undefined' && (window as any).applyFormatOnEmptyCell) {
              (window as any).applyFormatOnEmptyCell(cell, false);
            }
          }
        });
      // }
      this.quill.scrollSelectionIntoView();
      
      // After column insertion, just refresh the table selection state
      // Don't try to re-select specific cells as DOM references may be stale
      setTimeout(() => {
        try {
          // Clear any existing cell selections and refresh table state
          this.tableBetter.cellSelection.clearSelected();
          
          // Refresh cell selection to ensure we have valid references after the DOM mutation
          this.refreshCellSelection();
        } catch (error) {
          console.warn('Error refreshing selection after column insertion:', error);
        }
      }, 50);
    } catch (error) {
      console.error('Error inserting column:', error);
    }
  }

  insertParagraph(offset: number) {
    const blot = Quill.find(this.table) as TableContainer;
    const index = this.quill.getIndex(blot);
    const length = offset > 0 ? blot.length() : 0;

    
    const delta = new Delta()
      .retain(index + length)
      .insert('\n');
    this.quill.updateContents(delta, Quill.sources.USER);
    this.quill.setSelection(index + length, Quill.sources.SILENT);
   
    // Apply default font and size formatting
    const tableModule = this.quill.getModule('table') as Table;
        
    // Get default formats from the table module
    const defaultFont = tableModule?.defaultFormats?.font || tableModule?.lastFormat?.font || 'Helvetica';
    const defaultSize = tableModule?.defaultFormats?.size || tableModule?.lastFormat?.size || '18pt';

 
       // Apply default formats at current cursor position
       // Use silent source for the first formatting to avoid multiple updates
       this.quill.format('font', defaultFont, Quill.sources.SILENT);
       this.quill.format('size', defaultSize, Quill.sources.USER); // USER source for the last one to trigger an update
 
    // Apply to a small range to ensure formats stick
    const range = this.quill.getSelection() || { index: index, length: 1 };
    this.quill.formatText(range.index, Math.max(1, range.length), {
      font: defaultFont,
      size: defaultSize
    }, Quill.sources.USER);

    const fontPickerLabel = document.querySelector('.ql-font .ql-picker-label');
    const sizePickerLabel = document.querySelector('.ql-size .ql-picker-label');
    if (fontPickerLabel) {
      fontPickerLabel.classList.add("ql-active");
      fontPickerLabel.setAttribute('data-value', `${defaultFont}`);
    }
    if (sizePickerLabel) {
      sizePickerLabel.classList.add("ql-active");
      sizePickerLabel.setAttribute('data-value', `${defaultSize}`);
    }
  }

  insertRow(td: HTMLElement, offset: number) {
    try {
      // Validate td parameter
      if (!td) {
        console.warn('Cannot insert row: td parameter is null or undefined');
        return;
      }

      // Check if td is still connected to the DOM
      if (!td.isConnected) {
        console.warn('Cannot insert row: td element is disconnected from DOM');
        // Try to reconnect the cell reference
        const reconnectedTd = this.reconnectCellReference(td);
        if (reconnectedTd) {
          console.log('Reconnected td reference for row insertion');
          td = reconnectedTd;
        } else {
          console.error('Cannot insert row: unable to reconnect td reference');
          return;
        }
      }
      
      // Find the tdBlot
      const tdBlot = Quill.find(td) as TableCell;
      if (!tdBlot) {
        console.warn('Cannot insert row: tdBlot not found');
        return;
      }
      
      // Get row offset
      const index = tdBlot.rowOffset();
      if (index === null || index === undefined) {
        console.warn('Cannot insert row: invalid row index');
        return;
      }
      
      // Get table blot
      const tableBlot = tdBlot.table();
      if (!tableBlot) {
        console.warn('Cannot insert row: tableBlot not found');
        return;
      }
      
      // Check if tdBlot.statics exists
      if (!tdBlot.statics) {
        console.warn('Cannot insert row: tdBlot.statics is undefined');
        return;
      }
      
      const isTh = tdBlot.statics.blotName === TableTh.blotName;
      
      // Track existing cells before insertion to identify newly created ones
      const existingCells = new Set(Array.from(this.table?.querySelectorAll('td, th') || []));
      
      if (offset > 0) {
        // Safely get rowspan attribute
        let rowspan = 1;
        try {
          rowspan = td.getAttribute('rowspan') ? parseInt(td.getAttribute('rowspan'), 10) || 1 : 1;
        } catch (error) {
          console.warn('Error parsing rowspan attribute:', error);
        }
        
        tableBlot.insertRow(index + offset + rowspan - 1, offset, isTh);
      } else {
        tableBlot.insertRow(index + offset, offset, isTh);
      }
      console.log()
      
      // Apply formats to ALL newly inserted cells in the row
      // This ensures consistent formatting across all new cells
      const allCellsAfterInsertion = Array.from(this.table?.querySelectorAll('td, th') || []);
      const newCells = allCellsAfterInsertion.filter(cell => !existingCells.has(cell));
      
      console.log('🔍 INSERT-ROW: Applying formats to', newCells.length, 'new cells');
      
      // Simply call applyFormatOnEmptyCell for each new cell
      // This ensures all cells get the current active formatting consistently
      newCells.forEach((cell: any, index: number) => {
        console.log(`🔍 INSERT-ROW: Applying format to new cell ${index + 1}/${newCells.length}`, cell);
        
        if (typeof window !== 'undefined' && (window as any).applyFormatOnEmptyCell) {
          try {
            (window as any).applyFormatOnEmptyCell(cell, false);
            console.log(`🔍 INSERT-ROW: Successfully applied format to cell ${index + 1}`);
          } catch (error) {
            console.error(`🔍 INSERT-ROW: Failed to apply format to cell ${index + 1}:`, error);
          }
        } else {
          console.warn('🔍 INSERT-ROW: applyFormatOnEmptyCell not available');
        }
      });
      
      this.quill.scrollSelectionIntoView();
      
      // After row insertion, refresh the cell selection to ensure we have valid references
      // Use setTimeout to allow DOM to update before updating selection
      setTimeout(() => {
        try {
          // Clear any existing cell selections and refresh table state
          this.tableBetter.cellSelection.clearSelected();
          
          // Refresh cell selection to ensure we have valid references after the DOM mutation
          this.refreshCellSelection();
        } catch (error) {
          console.warn('Error updating selection after row insertion:', error);
        }
      }, 50);
    } catch (error) {
      console.error('Error inserting row:', error);
    }
  }

  mergeCells() {
    const { selectedTds } = this.tableBetter.cellSelection;
    const { computeBounds, leftTd } = this.getSelectedTdsInfo();
    const leftTdBlot = Quill.find(leftTd) as TableCell;
    const [formats, cellId] = getCellFormats(leftTdBlot);
    const head = leftTdBlot.children.head;
    const tableBlot = leftTdBlot.table();
    const rows = tableBlot.tbody().children as LinkedList<TableRow>;
    const row = leftTdBlot.row();
    const colspan = row.children.reduce((colspan: number, td: TableCell) => {
      const tdCorrectBounds = getCorrectBounds(td.domNode, this.quill.container);
      if (
        tdCorrectBounds.left >= computeBounds.left &&
        tdCorrectBounds.right <= computeBounds.right
      ) {
        colspan += ~~td.domNode.getAttribute('colspan') || 1;
      }
      return colspan;
    }, 0);
    const rowspan = rows.reduce((rowspan: number, row: TableRow) => {
      const rowCorrectBounds = getCorrectBounds(row.domNode, this.quill.container);
      if (
        rowCorrectBounds.top >= computeBounds.top &&
        rowCorrectBounds.bottom <= computeBounds.bottom
      ) {
        let minRowspan = Number.MAX_VALUE;
        row.children.forEach((td: TableCell) => {
          const rowspan = ~~td.domNode.getAttribute('rowspan') || 1;
          minRowspan = Math.min(minRowspan, rowspan);
        });
        rowspan += minRowspan;
      }
      return rowspan;
    }, 0);
    let offset = 0;
    for (const td of selectedTds) {
      if (leftTd.isEqualNode(td)) continue;
      const blot = Quill.find(td) as TableCell;
      blot.moveChildren(leftTdBlot);
      blot.remove();
      if (!blot.parent?.children?.length) offset++;
    }
    if (offset) {
      // Subtract the number of rows deleted by the merge
      row.children.forEach((child: TableCell) => {
        if (child.domNode.isEqualNode(leftTd)) return;
        const rowspan = child.domNode.getAttribute('rowspan');
        const [formats] = getCellFormats(child);
        // @ts-expect-error
        child.replaceWith(child.statics.blotName, { ...formats, rowspan: rowspan - offset });
      });
    }
    leftTdBlot.setChildrenId(cellId);
    // @ts-expect-error
    head.format(leftTdBlot.statics.blotName, { ...formats, colspan, rowspan: rowspan - offset });
    this.tableBetter.cellSelection.setSelected(head.parent.domNode);
    this.quill.scrollSelectionIntoView();
  }

  selectColumn() {
    const { computeBounds, leftTd, rightTd } = this.getSelectedTdsInfo();
    const selectTds = getComputeSelectedTds(computeBounds, this.table, this.quill.container, 'column');
    const { selTds } = this.getCorrectTds(selectTds, computeBounds, leftTd, rightTd);
    this.tableBetter.cellSelection.setSelectedTds(selTds);
  }

  selectRow() {
    const rows = this.getCorrectRows();
    const selectTds = rows.reduce((selTds: Element[], row: TableRow) => {
      selTds.push(...Array.from(row.domNode.children));
      return selTds;
    }, []);
    this.tableBetter.cellSelection.setSelectedTds(selectTds);
  }

  setCellsMap(cell: TableCell, map: TableCellMap) {
    const key: string = cell.domNode.getAttribute('data-row');
    if (map.has(key)) {
      map.set(key, [...map.get(key), cell.domNode]);
    } else {
      map.set(key, [cell.domNode]);
    }
  }

  showMenus() {
    // Just remove the class — updateMenus() will re-apply inline position styles
    this.root.classList.remove('ql-hidden');

    this.lastMenuShowTime = Date.now();
    console.log('🔍 SHOWMENUS: Menu shown at', this.lastMenuShowTime);

    if (typeof window !== 'undefined' && /iPad|iPhone|iPod/.test(navigator.userAgent)) {
      const lastClickedCell = this.tableBetter.lastClickedCell;
      if (lastClickedCell && lastClickedCell.isConnected &&
          this.tableBetter.cellSelection.selectedTds.length === 0) {
        console.log('showMenus: iOS logic - restoring cell selection');
      }
    }
  }

  splitCell() {
    const { selectedTds } = this.tableBetter.cellSelection;
    const { leftTd } = this.getSelectedTdsInfo();
    const leftTdBlot = Quill.find(leftTd) as TableCell;
    const head = leftTdBlot.children.head;
    for (const td of selectedTds) {
      const colspan = ~~td.getAttribute('colspan') || 1;
      const rowspan = ~~td.getAttribute('rowspan') || 1;
      if (colspan === 1 && rowspan === 1) continue;
      const columnCells: [TableRow, string, TableCell | null][] = [];
      const { width, right } = td.getBoundingClientRect();
      const blot = Quill.find(td) as TableCell;
      const tableBlot = blot.table();
      const nextBlot = blot.next;
      const rowBlot = blot.row();
      if (rowspan > 1) {
        if (colspan > 1) {
          let nextRowBlot = rowBlot.next;
          for (let i = 1; i < rowspan; i++) {
            const { ref, id } = this.getRefInfo(nextRowBlot, right);
            for (let j = 0; j < colspan; j++) {
              columnCells.push([nextRowBlot, id, ref]);
            }
            nextRowBlot && (nextRowBlot = nextRowBlot.next);
          }
        } else {
          let nextRowBlot = rowBlot.next;
          for (let i = 1; i < rowspan; i++) {
            const { ref, id } = this.getRefInfo(nextRowBlot, right);
            columnCells.push([nextRowBlot, id, ref]);
            nextRowBlot && (nextRowBlot = nextRowBlot.next);
          }
        }
      }
      if (colspan > 1) {
        const id = td.getAttribute('data-row');
        for (let i = 1; i < colspan; i++) {
          columnCells.push([rowBlot, id, nextBlot]);
        }
      }
      for (const [row, id, ref] of columnCells) {
        tableBlot.insertColumnCell(row, id, ref);
      }
      const [formats] = getCellFormats(blot);
      blot.replaceWith(blot.statics.blotName, {
        ...formats,
        width: ~~(width / colspan),
        colspan: null,
        rowspan: null
      });
    }
    this.tableBetter.cellSelection.setSelected(head.parent.domNode);
    this.quill.scrollSelectionIntoView();
  }

  toggleAttribute(list: HTMLUListElement, tooltip: HTMLDivElement, e?: PointerEvent) {
    // @ts-expect-error
    if (e && e.target.closest('li.ql-table-header-row')) return;
    if (this.prevList && !this.prevList.isEqualNode(list)) {
      this.prevList.classList.add('ql-hidden');
      this.prevTooltip.classList.remove('ql-table-tooltip-hidden');
    }
    if (!list) return;
    
    // Check if we're showing a dropdown (not hiding)
    const isShowing = list.classList.contains('ql-hidden');
    
    // Flag to track whether the table menu is currently open
    // This is used to prevent the menu from being closed when the user interacts with the table
    // while the menu is still open
    if (isShowing) {
      (window as any).isTableMenuOpen = true;
      // REMOVED: this.quill.blur() - This was clearing the selection!
      
      list.classList.remove('ql-hidden');
      tooltip.classList.add('ql-table-tooltip-hidden');
      this.prevList = list;
      this.prevTooltip = tooltip;
      this.isDropdownOpen = true;
      
      // Ensure cell selection is maintained
      if (this.tableBetter.cellSelection.selectedTds.length > 0) {
        const selectedCell = this.tableBetter.cellSelection.selectedTds[0];
        if (selectedCell && selectedCell.isConnected) {
          // Re-apply the focused class to ensure blue selection is visible
          selectedCell.classList.add('ql-cell-focused');
        }
      }
    } else {
      (window as any).isTableMenuOpen = false; // Clear the flag when the menu closes to allow other interactions
      // When hiding, restore focus to the previously selected cell
      list.classList.add('ql-hidden');
      tooltip.classList.remove('ql-table-tooltip-hidden');
      this.prevList = list;
      this.prevTooltip = tooltip;
      this.isDropdownOpen = false;
    }
  }

  toggleHeaderRow() {
    const { selectedTds, hasTdTh } = this.tableBetter.cellSelection;
    const { hasTd, hasTh } = hasTdTh(selectedTds);
    if (!hasTd && hasTh) {
      this.convertToRow();
    } else {
      this.convertToHeaderRow();
    }
  }

  toggleHeaderRowSwitch(value?: string) {
    if (!this.tableHeaderRow) return;
    const switchInner = this.tableHeaderRow.querySelector('.ql-table-switch-inner');
    if (!value) {
      const ariaChecked = switchInner.getAttribute('aria-checked');
      value = ariaChecked === 'false' ? 'true' : 'false';
    }
    switchInner.setAttribute('aria-checked', value);
  }

  updateMenus(table: HTMLElement = this.table, clickedCell?: HTMLElement) {
    if (!table) return;

    // Check if any dropdown list is currently visible
    // If so, don't update menu position to avoid repositioning
    const openDropdowns = this.root.querySelectorAll('.ql-table-dropdown-list:not(.ql-hidden)');
    if (openDropdowns.length > 0) {
      return;
    }

    try {
      this.root.classList.remove('ql-table-triangle-none');
      const [tableBounds, containerBounds] = this.getCorrectBounds(table);
      const { left, right, top, bottom } = tableBounds;
      const { height, width } = this.root.getBoundingClientRect();
      const toolbar = this.quill.getModule('toolbar');
      // @ts-expect-error
      const computedStyle = getComputedStyle(toolbar.container);

      // Use the clicked cell if provided, otherwise use selected cells
      const cellToPosition = clickedCell || (this.tableBetter.cellSelection.selectedTds.length > 0 ? this.tableBetter.cellSelection.selectedTds[0] : null);

      let correctTop: number;
      let correctLeft: number;

      // Check if we have a cell to position the menu relative to it
      if (cellToPosition && cellToPosition.isConnected) {
        // CRITICAL FIX: Use viewport-relative coordinates with getBoundingClientRect()
        // This ensures menu is positioned correctly relative to the visible viewport
        // Never use document-relative coordinates which cause menu to float mid-table
        const cellRect = (cellToPosition as HTMLElement).getBoundingClientRect();
        const menuHeight = height;
        const menuWidth = width;
        const safeGap = 6;
        const viewportHeight = window.innerHeight;
        const viewportWidth = window.innerWidth;
        
        // CRITICAL FIX: Use scrollend event to guarantee positioning ONLY after scroll completes
        // This prevents stale coordinates from being used during scroll animation
        // Menu is invisible during scroll, then repositioned and shown only after scroll ends
        
        // Hide menu instantly before any scroll
        this.root.classList.add('ql-hidden');
        this.root.style.visibility = 'hidden';
        this.root.style.opacity = '0';
        console.log('🔍 UPDATEMENUS: Menu hidden before positioning');
        
        // Define the repositioning logic that will run after scroll completes
        const repositionMenuAfterScroll = () => {
          console.log('🔍 UPDATEMENUS: Scroll ended, repositioning menu with fresh coordinates');
          
          // CRITICAL FIX: Make menu renderable but invisible to get real dimensions
          // Hidden elements return offsetWidth = 0, causing wrong calculations
          // Remove ql-hidden class but keep visibility hidden and opacity 0
          this.root.classList.remove('ql-hidden');
          this.root.style.visibility = 'hidden';
          this.root.style.opacity = '0';
          this.root.style.display = 'flex'; // Ensure menu is rendered in DOM
          
          // CRITICAL FIX: Use getBoundingClientRect() for measurement, not offsetHeight
          // offsetHeight returns 0 when element is hidden, but getBoundingClientRect() works
          // even with visibility:hidden as long as display is not 'none'
          // Position menu off-screen during measurement to avoid flash
          this.root.style.top = '-9999px';
          this.root.style.left = '-9999px';
          
          // NOW measure using getBoundingClientRect() - returns real dimensions
          const menuRect = this.root.getBoundingClientRect();
          const freshMenuHeight = menuRect.height || 44;
          const freshMenuWidth = menuRect.width || 280;
          console.log('🔍 UPDATEMENUS: Menu dimensions measured (height:', freshMenuHeight, 'width:', freshMenuWidth, ')');
          
          // Get FRESH cell coordinates AFTER scroll has fully completed
          const freshCellRect = (cellToPosition as HTMLElement).getBoundingClientRect();
          console.log('🔍 UPDATEMENUS: Cell bounds (top:', freshCellRect.top, 'bottom:', freshCellRect.bottom, 'left:', freshCellRect.left, ')');
          
          let menuTop: number;
          let menuLeft: number;
          
          // CRITICAL: Calculate space needed (full menu height + gap)
          const safeGap = 8;
          const spaceNeeded = freshMenuHeight + safeGap;
          
          // STEP 1: Position menu above or below the TAPPED CELL
          // Menu must follow the individual cell, not the table
          if (freshCellRect.top >= spaceNeeded) {
            // Enough space above cell - show menu above cell
            menuTop = freshCellRect.top - freshMenuHeight - safeGap;
            this.root.classList.add('ql-table-triangle-up');
            this.root.classList.remove('ql-table-triangle-down');
            this.root.classList.remove('ql-table-triangle-none');
            console.log('🔍 UPDATEMENUS: Menu positioned ABOVE cell (cell top:', freshCellRect.top, 'needed:', spaceNeeded, ')');
          } else {
            // Not enough space above cell - show menu below cell
            menuTop = freshCellRect.bottom + safeGap;
            this.root.classList.add('ql-table-triangle-down');
            this.root.classList.remove('ql-table-triangle-up');
            this.root.classList.remove('ql-table-triangle-none');
            console.log('🔍 UPDATEMENUS: Menu positioned BELOW cell (cell bottom:', freshCellRect.bottom, ')');
          }
          
          // STEP 2: Clamp vertical position to viewport bounds
          menuTop = Math.max(
            safeGap,
            Math.min(menuTop, viewportHeight - freshMenuHeight - safeGap)
          );
          
          // STEP 3: Align menu horizontally with cell left edge, clamp to viewport
          menuLeft = freshCellRect.left;
          if (menuLeft + freshMenuWidth > viewportWidth - safeGap) {
            menuLeft = viewportWidth - freshMenuWidth - safeGap;
          }
          menuLeft = Math.max(safeGap, menuLeft);
          console.log('🔍 UPDATEMENUS: Menu horizontal position (left:', menuLeft, 'width:', freshMenuWidth, 'viewport:', viewportWidth, ')');
          
          // STEP 3: Apply position FIRST, then show menu
          // This ensures user never sees menu jump or flash at wrong position
          correctTop = menuTop;
          correctLeft = menuLeft;
          
          setElementProperty(this.root, {
            left: `${correctLeft}px`,
            top: `${correctTop}px`
          });
          
          // NOW show menu - position is already set correctly
          this.root.style.position = 'fixed';
          this.root.style.visibility = 'visible';
          this.root.style.opacity = '1';
          console.log('🔍 UPDATEMENUS: Menu shown at fixed position (top:', correctTop, 'left:', correctLeft, ')');
        };
        
        // Check if scrollend event is supported
        if ('onscrollend' in window) {
          console.log('🔍 UPDATEMENUS: Using scrollend event for positioning');
          
          // Set flag to prevent scroll listener from hiding menu
          this.isProgrammaticScroll = true;
          
          // CRITICAL FIX: Safety timeout to prevent flag from being stuck
          // If scroll takes longer than 2s, force reset flag
          const safetyTimeout = (window as any).setTimeout(() => {
            console.log('🔍 UPDATEMENUS: Safety timeout triggered - resetting isProgrammaticScroll');
            this.isProgrammaticScroll = false;
          }, 2000);
          
          // Trigger scroll if needed
          (cellToPosition as HTMLElement).scrollIntoView({ behavior: 'smooth', block: 'center' });
          
          // Wait for scrollend event - guaranteed to fire after scroll completes
          const onScrollEnd = () => {
            window.removeEventListener('scrollend', onScrollEnd);
            clearTimeout(safetyTimeout); // Cancel safety timeout
            this.isProgrammaticScroll = false;
            repositionMenuAfterScroll();
          };
          
          window.addEventListener('scrollend', onScrollEnd);
        } else {
          console.log('🔍 UPDATEMENUS: scrollend not supported, using debounced scroll event');
          
          // Fallback for browsers without scrollend support (older iOS WebView)
          this.isProgrammaticScroll = true;
          
          // CRITICAL FIX: Safety timeout to prevent flag from being stuck
          // If scroll takes longer than 2s, force reset flag
          const safetyTimeout = (window as any).setTimeout(() => {
            console.log('🔍 UPDATEMENUS: Safety timeout triggered - resetting isProgrammaticScroll');
            this.isProgrammaticScroll = false;
          }, 2000);
          
          (cellToPosition as HTMLElement).scrollIntoView({ behavior: 'smooth', block: 'center' });
          
          let scrollEndTimer: number;
          const onScroll = () => {
            clearTimeout(scrollEndTimer);
            scrollEndTimer = window.setTimeout(() => {
              // Scroll has been idle 150ms - treat as ended
              window.removeEventListener('scroll', onScroll);
              clearTimeout(safetyTimeout); // Cancel safety timeout
              this.isProgrammaticScroll = false;
              repositionMenuAfterScroll();
            }, 150);
          };
          
          (window as any).addEventListener('scroll', onScroll);
        }
        
        return;
      } else {
        // Fall back to table-based positioning if no cell is selected
        correctTop = top - height - 10;
        correctLeft = (left + right - width) >> 1;

        if (correctTop > -parseInt(computedStyle.paddingBottom)) {
          this.root.classList.add('ql-table-triangle-up');
          this.root.classList.remove('ql-table-triangle-down');
        } else {
          if (bottom > containerBounds.height) {
            correctTop = containerBounds.height + 10;
          } else {
            correctTop = bottom + 10;
          }
          this.root.classList.add('ql-table-triangle-down');
          this.root.classList.remove('ql-table-triangle-up');
        }
        
        // Ensure the menu stays within the container bounds
        if (correctLeft < containerBounds.left) {
          correctLeft = 0;
          this.root.classList.add('ql-table-triangle-none');
        } else if (correctLeft + width > containerBounds.right) {
          correctLeft = containerBounds.right - width;
          this.root.classList.add('ql-table-triangle-none');
        }
      }

      // Apply the position synchronously (no requestAnimationFrame to prevent DOM mutations)
      setElementProperty(this.root, {
        left: `${correctLeft}px`,
        top: `${correctTop}px`
      });
    } catch (error) {
      console.error('Error in updateMenus:', error);
    }
  }

  updateScroll(scroll: boolean) {
    this.scroll = scroll;
  }

  updateTable(table: HTMLElement) {
    this.table = table;
  }

  reconnectCellReference(disconnectedTd: HTMLElement): HTMLElement | null {
    if (!this.table || !disconnectedTd) {
      return null;
    }

    try {
      // Get all cells in the current table
      const allCells = Array.from(this.table.querySelectorAll('td, th'));
      
      // Try to find a cell with the same data-row attribute first
      const originalRowId = disconnectedTd.getAttribute('data-row');
      if (originalRowId) {
        const matchingRowCells = allCells.filter(cell => 
          cell.getAttribute('data-row') === originalRowId && cell.isConnected
        );
        if (matchingRowCells.length > 0) {
          return matchingRowCells[0] as HTMLElement;
        }
      }
      
      // For newly inserted cells that might not have data-row attributes yet,
      // try to find a cell in the same approximate position
      const disconnectedRect = disconnectedTd.getBoundingClientRect();
      if (disconnectedRect.width > 0 && disconnectedRect.height > 0) {
        let closestCell: HTMLElement | null = null;
        let minDistance = Infinity;
        
        for (const cell of allCells) {
          if (!cell.isConnected) continue;
          
          const cellRect = cell.getBoundingClientRect();
          const distance = Math.sqrt(
            Math.pow(cellRect.left - disconnectedRect.left, 2) + 
            Math.pow(cellRect.top - disconnectedRect.top, 2)
          );
          
          if (distance < minDistance) {
            minDistance = distance;
            closestCell = cell as HTMLElement;
          }
        }
        
        // Only use the closest cell if it's reasonably close (within 50px)
        if (closestCell && minDistance < 50) {
          return closestCell;
        }
      }
      
      // Final fallback: return any connected cell
      const connectedCells = allCells.filter(cell => cell.isConnected);
      if (connectedCells.length > 0) {
        return connectedCells[0] as HTMLElement;
      }
      
      return null;
    } catch (error) {
      console.warn('Error in reconnectCellReference:', error);
      return null;
    }
  }
}

export default TableMenus;