import { Component, OnInit, OnDestroy } from '@angular/core';
import { CommonModule, DatePipe } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { LayoutComponent } from '../../layout/layout/layout';
import { ToastrService } from 'ngx-toastr';
import { Service } from '../../dashboard/service';
import { AgGridAngular } from 'ag-grid-angular';
import { StatusFilterComponent } from '../../status-filter';
import { DateFilterComponent } from '../../date-filter'; 
import {
  ColDef,
  GridApi,
  GridReadyEvent,
  IGetRowsParams,
  IDatasource,
  ModuleRegistry,
  AllCommunityModule,
  InfiniteRowModelModule,
  RowModelType
} from 'ag-grid-community';
import { Subject, Subscription } from 'rxjs';
import { debounceTime, distinctUntilChanged } from 'rxjs/operators';

// Registering here too (not just in dashboard.ts) since this page can be
// the first — or only — route ag-grid ever mounts on in a given session
// if the user navigates straight to /admin. registerModules() is safe to
// call more than once.
ModuleRegistry.registerModules([AllCommunityModule, InfiniteRowModelModule]);

interface AdminRole {
  id: number;
  role_name: string;
}

// Normalizes whatever shape the backend sends a role in (id/role_id,
// role_name/name) so the template never has to guess. Still used for the
// "Existing Roles" side panel, which hits a separate, non-paginated
// /roles endpoint — unrelated to the users grid below.
function normalizeRole(r: any): AdminRole {
  return { id: r.id ?? r.role_id, role_name: r.role_name ?? r.name ?? 'Unnamed role' };
}

type ModalMode = 'configuration' | 'role' | null;

@Component({
  selector: 'app-user',
  standalone: true,
  imports: [CommonModule, FormsModule, LayoutComponent, AgGridAngular],
  templateUrl: './user.html',
  styleUrl: './user.css',
  providers: [DatePipe],
})
export class User implements OnInit, OnDestroy {
  adminName = sessionStorage.getItem('userName') || '';

  roles: AdminRole[] = [];
  loadingRoles = false;
  submitting = false;

  constructor(
    private service: Service,
    private toastr: ToastrService,
    private datePipe: DatePipe,
  ) {}

  // ── AG Grid — users table ─────────────────────────────
  // Bound to POST /users/list, same infinite-row-model + server-side
  // filter/sort/quickSearch pattern as the Concept Dashboard grid (see
  // dashboard.ts). Columns are built dynamically from the columnMetadata
  // the backend returns, plus one client-only "Action" column appended
  // at the end for the Modify button.
  rowModelType: RowModelType = 'infinite';
  PAGE_SIZE = 20;
  cacheOverflowSize = 2;
  maxConcurrentDatasourceRequests = 2;
  infiniteInitialRowCount = 1;
  maxBlocksInCache = 2;
  pagination = true;
  paginationPageSizeSelector = [20, 50, 100];

  loading = false;
  initialLoad = true;
  gridApi!: GridApi;
  private isGridReady = false;

  columnDefs: ColDef[] = [];

  private readonly ACTION_COL_ID = 'rowAction';

  defaultColDef: ColDef = {
    sortable: true,
    filter: true,
    resizable: true,
    suppressMovable: true,
    minWidth: 120,
    maxWidth: 225,
    filterParams: {
      maxNumConditions: 1,
      suppressAndOrCondition: true,
      suppressConditionAndButton: true,
      buttons: ['reset', 'apply'],
      closeOnApply: true,
      filterOptions: [
        'contains',
        'notContains',
        'equals',
        'notEqual',
        'startsWith',
        'endsWith'
      ]
    },
  tooltipValueGetter: (params) => {
    return params.value ? params.value.toString() : '';
  }
  };

  overlayNoRowsTemplate = `
    <span class="custom-no-rows">
      ❌ No users found
    </span>
  `;
  overlayLoadingTemplate = `
  <div class="custom-loading-overlay">
    <div class="ag-custom-loading-cell" style="display: flex; align-items: center; gap: 6px; padding-left: 10px; line-height: 25px;">
      <img src="/Images/loader.gif" alt="Loading..." style="width: 150px; height: 150px;" />
      <span>Loading...</span>
    </div>
  </div>`;

  // ── Search ────────────────────────────────────────────
  // Same reasoning as dashboard.ts: [quickFilterText] only searches rows
  // already cached client-side under the infinite row model, so search
  // text is sent to the backend as its own `quickSearch` field on every
  // getRows() call instead.
  searchQuery = '';
  private searchSubject = new Subject<string>();
  private searchSub?: Subscription;

  private static readonly FILTER_STORAGE_KEY = 'admin_users_grid_filter_state_v1';

  private buildFilterModelForApi() {
    const raw = this.gridApi?.getFilterModel() || {};
    const finalModel: any = {};
    Object.keys(raw).forEach(colId => {
      const model = raw[colId];
      if (!model) return;
      if (model.filter == null && model.dateFrom == null && model.dateTo == null && model.values == null) return;
      finalModel[colId] = model;
    });
    return finalModel;
  }

  private loadSavedFilterState(): { search: string; filterModel: any } | null {
    try {
      const raw = sessionStorage.getItem(User.FILTER_STORAGE_KEY);
      return raw ? JSON.parse(raw) : null;
    } catch {
      return null;
    }
  }

  private persistFilterState(): void {
    try {
      const state = {
        search: this.searchQuery,
        filterModel: this.gridApi ? this.gridApi.getFilterModel() : null
      };
      sessionStorage.setItem(User.FILTER_STORAGE_KEY, JSON.stringify(state));
    } catch {
      // sessionStorage unavailable — filters just won't persist, nothing else breaks.
    }
  }

  onFilterChanged(): void {
    this.persistFilterState();
  }

  onSearchChange(): void {
    this.searchSubject.next(this.searchQuery);
  }

  // ── Lifecycle ─────────────────────────────────────────
  ngOnInit(): void {
    this.loadRoles();

    const saved = this.loadSavedFilterState();
    if (saved?.search) {
      this.searchQuery = saved.search;
    }

    this.searchSub = this.searchSubject.pipe(
      debounceTime(300),
      distinctUntilChanged()
    ).subscribe(() => {
      this.persistFilterState();
      this.refreshGridData();
    });
  }

  ngOnDestroy(): void {
    this.searchSub?.unsubscribe();
    this.closeRolesPopover();
  }

  private refreshGridData(): void {
    if (!this.gridApi) return;
    const api: any = this.gridApi;
    if (typeof api.refreshInfiniteCache === 'function') {
      api.refreshInfiniteCache();
    } else if (typeof api.purgeInfiniteCache === 'function') {
      api.purgeInfiniteCache();
    } else {
      this.setDatasource();
    }
  }

  onGridReady(event: GridReadyEvent): void {
    this.gridApi = event.api;
    this.isGridReady = true;

    const saved = this.loadSavedFilterState();
    const hasSavedFilter = !!(saved?.filterModel && Object.keys(saved.filterModel).length > 0);

    if (hasSavedFilter && this.columnDefs.length === 0) {
      this.bootstrapColumnsThenStart();
    } else if (hasSavedFilter) {
      this.applySavedFilterThenStart();
    } else {
      this.setDatasource();
    }
  }

  private bootstrapColumnsThenStart(): void {
    this.loading = true;

    this.service.getUsers(this.buildRequestBody(1, 1, [])).subscribe({
      next: (res: any) => {
        this.columnDefs = this.buildColumnDefsFromMetadata(res.columnMetadata);
        this.gridApi.setGridOption('columnDefs', this.columnDefs);
        this.applySavedFilterThenStart();
      },
      error: () => {
        this.loading = false;
        this.initialLoad = false;
        console.error('Failed to load user columns.');
        this.setDatasource();
      }
    });
  }

  private applySavedFilterThenStart(): void {
    const saved = this.loadSavedFilterState();

    if (saved?.filterModel && Object.keys(saved.filterModel).length > 0) {
      const result: any = this.gridApi.setFilterModel(saved.filterModel);
      if (result && typeof result.then === 'function') {
        result.then(() => this.setDatasource());
      } else {
        this.setDatasource();
      }
    } else {
      this.setDatasource();
    }
  }

  private buildRequestBody(page: number, pageSize: number, sortModel: any[]) {
    return {
      page,
      page_size: pageSize,
      sortModel,
      filterModel: this.buildFilterModelForApi(),
      quickSearch: this.searchQuery?.trim() || ''
    };
  }

  setDatasource(): void {
    this.loading = true;

    const datasource: IDatasource = {
      getRows: (params: IGetRowsParams) => {
        const startRow = params.startRow;
        const currentPage = Math.floor(startRow / this.PAGE_SIZE) + 1;

        this.loading = true;

        if (!this.initialLoad) {
          this.gridApi.setGridOption('loading', true);
          this.gridApi.showLoadingOverlay();
        }

        const requestBody = this.buildRequestBody(currentPage, this.PAGE_SIZE, params.sortModel);

        this.service.getUsers(requestBody).subscribe({
          next: (res: any) => {
            if (res?.columnMetadata && this.columnDefs.length === 0) {
              this.columnDefs = this.buildColumnDefsFromMetadata(res.columnMetadata);
              this.gridApi.setGridOption('columnDefs', this.columnDefs);
            }

            const rows = res?.data ?? [];
            const totalCount = res?.totalCount ?? 0;
            params.successCallback(rows, totalCount);
            this.gridApi.setGridOption('loading', false);
            this.loading = false;
            this.initialLoad = false;

            if (rows.length === 0) {
              this.gridApi.showNoRowsOverlay();
            } else {
              this.gridApi.hideOverlay();
            }
          },
          error: () => {
            this.loading = false;
            this.initialLoad = false;
            this.gridApi.hideOverlay();
            this.gridApi.setGridOption('loading', false);
            params.failCallback();
            console.error('Failed to load users.');
            this.toastr.error('Something went wrong Connecting to the server.', 'Error');
          }
        });
      }
    };

    this.gridApi.setGridOption('datasource', datasource);
  }

  /** Builds grid columns from the backend's columnMetadata, same as
   *  dashboard.ts's equivalent, with three users-specific tweaks:
   *   - `roles` gets rendered as pill badges instead of the raw
   *     comma-separated string the backend sends it as.
   *   - `is_active` gets rendered as the colored status dot the old
   *     plain-HTML table used, instead of a bare true/false cell.
   *   - a client-only "Action" column (Modify button) is appended at
   *     the end — it isn't part of columnMetadata since the backend
   *     doesn't know anything about frontend row actions.
   *
   *  ASSUMES the backend row/column keys are `id`, `username`, `name`,
   *  `roles`, `is_active`, `created_at`, `modified_at` (matching
   *  um.fetch_all_users' presumed dict shape). If the real column keys
   *  differ, update the field checks below (`field === 'roles'` etc.)
   *  to match. */
  private buildColumnDefsFromMetadata(metadata: any): ColDef[] {
    // modified_at isn't part of the target table layout (System
    // Administration mockup shows ID / User Name / Name / Assigned To /
    // Created Date / State / Action only) — drop it from the grid. It's
    // still returned by the backend and still filterable/sortable
    // elsewhere if ever needed; this just keeps it off this view.
    // const orderedFields: string[] = Object.keys(metadata || {}).filter(
    //   (f) => f !== 'modified_at' && f !== 'modifiedAt'
    // );
    const orderedFields: string[] = Object.keys(metadata || {});

    const dataColumns: ColDef[] = orderedFields.map((field: string) => {
      const meta = metadata[field] || {};
      const isDateColumn = meta.type === 'date';

      const filter =
        meta.type === 'date' ? DateFilterComponent :
        meta.type === 'set' ? StatusFilterComponent :
        meta.type === 'number' || meta.type === 'float' ? 'agNumberColumnFilter' :
        'agTextColumnFilter';

      const filterParams = isDateColumn
        ? {
            filterOptions: meta.filterOperations || [
              'equals', 'notEqual', 'lessThan', 'greaterThan', 'inRange'
            ]
          }
        : meta.type === 'set'
        ? { values: meta.options || [] }
        : {
            ...this.defaultColDef.filterParams,
            filterOptions: meta.filterOperations
          };

      const column: ColDef = {
        field,
        colId: field,
        headerName: field.replace(/_/g, ' ').replace(/\b\w/g, (c) => c.toUpperCase()),
        sortable: true,
        resizable: true,
        suppressMovable: true,
        minWidth: 120,
        flex: 1,
        filter,
        filterParams,
        valueGetter: (p: { data: { [x: string]: any } }) => p.data?.[field]
      };

      if (field === 'id') {
        // Just a short numeric id — doesn't need to stretch and eat
        // space other columns actually need. flex: 0 stops it fighting
        // the other columns' flex: 1 for extra width.
        column.headerName = 'ID';
        column.minWidth = 40;
        column.maxWidth = 60;
        column.flex = 0;
      }
      if (field ==='username') {
        column.headerName = 'USERNAME';
        column.minWidth = 224;
        column.maxWidth = 230;
      }

      if (isDateColumn) {
        // Mockup shows date + time together ("11/02/2026  12:30 PM").
        column.valueFormatter = (params) => {
          return this.datePipe.transform(params.value, 'MM/dd/yyyy') ?? '';
        };
        column.minWidth = 190;
      }

      // Match by intent rather than one exact hardcoded key — the backend's
      // actual column key for "roles" / "is a user active" can (and does)
      // vary (roles / role_names / assigned_roles / userRoles, is_active /
      // active / status / state, etc). Matching on exact string equality
      // silently falls through to a plain untruncated-text column with no
      // badge/dot rendering, which is what was happening before this fix.
      const normalizedField = field.toLowerCase().replace(/[_\s]/g, '');
      const isRolesField = normalizedField.includes('role');
      const isActiveField =
        normalizedField.includes('active') ||
        normalizedField === 'status' ||
        normalizedField === 'state';

      if (isRolesField) {
        // 0 or 1 role renders as a plain badge. 2+ roles collapse into a
        // single compact "N roles" pill instead of listing every badge —
        // listing them all overflows the column unpredictably depending
        // on how long role names are, whereas a count badge keeps the
        // column width consistent no matter how many roles a user has.
        // The full list opens in a small panel on HOVER (see
        // scheduleRolesPopoverOpen/scheduleRolesPopoverClose) — a native
        // title tooltip is also kept as a fallback for touch/no-hover
        // input, where the mouseenter/mouseleave handlers never fire.
        column.headerName = 'Assigned To';
        column.minWidth = 200;
        column.cellClass = 'assigned-to-cell';
        // Explicitly disable the native tooltip for this column. Without this
        // it silently inherits defaultColDef's tooltipValueGetter (raw
        // comma-separated roles string), which renders as a native browser
        // tooltip behind/alongside the custom roles-popover below — the
        // "ghost list" showing through on hover.
        column.tooltipValueGetter = () => null;
        column.cellRenderer = (params: any) => {
          const names = this.rolesToNames(params.value);
          const container = document.createElement('div');
          container.className = 'role-badge-list';

          if (names.length === 0) {
            container.innerHTML = `<span class="admin-empty-inline">No roles</span>`;
            return container;
          }

          if (names.length === 1) {
            container.innerHTML = `<span class="role-badge">${this.escapeHtml(names[0])}</span>`;
            return container;
          }

          const countBadge = document.createElement('span');
          countBadge.className = 'role-badge role-count-badge';
          countBadge.textContent = `${names.length} roles`;
          // countBadge.title = names.join(', ');

          countBadge.addEventListener('mouseenter', () => {
            this.scheduleRolesPopoverOpen(countBadge, names);
          });
          countBadge.addEventListener('mouseleave', () => {
            this.scheduleRolesPopoverClose();
          });
          // Still openable by click/tap too, for touch devices with no hover.
          countBadge.addEventListener('click', (e) => {
            e.stopPropagation();
            this.toggleRolesPopover(countBadge, names);
          });

          container.appendChild(countBadge);
          return container;
        };
      }

      if (isActiveField) {
        column.headerName = 'State';
        // column.filter = false;
        column.sortable = true;
        column.minWidth = 110;
        column.maxWidth = 130;
        column.cellClass = 'state-cell';
        column.cellRenderer = (params: any) => {
          const raw = params.value;
          const active =
            raw === true || raw === 1 ||
            (typeof raw === 'string' && ['true', 'active', '1', 'yes'].includes(raw.toLowerCase()));
          const label = active ? 'Active' : 'Inactive';
          return `<span class="state-pill">
              <span class="state-dot ${active ? 'state-active' : 'state-inactive'}"></span>
              <span class="state-label ${active ? 'state-label-active' : 'state-label-inactive'}">${label}</span>
            </span>`;
        };
      }

      return column;
    });

    return [...dataColumns, this.buildActionColumn()];
  }

  /** Normalizes a roles cell value into a clean list of display names.
   *  Backend might send this as a comma-separated string ("Admin, QA"),
   *  an array of strings (["Admin", "QA"]), or an array of role objects
   *  ([{ role_name: "Admin" }, ...]) — handle all three so the badge
   *  renderer doesn't silently fail on a shape mismatch. */
  private rolesToNames(value: any): string[] {
    if (value == null) return [];
    const raw: any[] = Array.isArray(value) ? value : String(value).split(',');
    return raw
      .map((v) => (typeof v === 'string' ? v : v?.role_name ?? v?.name ?? ''))
      .map((s: string) => s.trim())
      .filter((s: string) => s.length > 0);
  }

  private escapeHtml(s: string): string {
    const div = document.createElement('div');
    div.textContent = s;
    return div.innerHTML;
  }

  // ── Roles "N roles" popover ─────────────────────────────
  // A single floating panel, appended to <body> (not the grid cell)
  // so it isn't clipped by the cell's overflow:hidden. Only one can be
  // open at a time. Primary trigger is hover (mouseenter/mouseleave on
  // the badge, plus mouseenter/mouseleave on the popover itself so
  // moving the pointer from badge to panel doesn't close it); a short
  // delay on close absorbs the small gap between the two elements.
  // Click/tap still works too (toggleRolesPopover), for touch input
  // that never fires hover events at all.
  private rolesPopoverEl: HTMLDivElement | null = null;
  private rolesPopoverAnchor: HTMLElement | null = null;
  private rolesPopoverCloseTimer: ReturnType<typeof setTimeout> | null = null;
  private readonly boundCloseRolesPopover = () => this.closeRolesPopover();

  private clearRolesPopoverCloseTimer(): void {
    if (this.rolesPopoverCloseTimer != null) {
      clearTimeout(this.rolesPopoverCloseTimer);
      this.rolesPopoverCloseTimer = null;
    }
  }

  private scheduleRolesPopoverOpen(anchor: HTMLElement, names: string[]): void {
    this.clearRolesPopoverCloseTimer();
    if (this.rolesPopoverAnchor === anchor) return; // already open for this badge
    this.openRolesPopover(anchor, names);
  }

  private scheduleRolesPopoverClose(): void {
    this.clearRolesPopoverCloseTimer();
    // Small grace period so moving the mouse from the badge onto the
    // popover itself (there's a few px gap between them) doesn't close
    // it before the popover's own mouseenter handler can cancel this.
    this.rolesPopoverCloseTimer = setTimeout(() => this.closeRolesPopover(), 150);
  }

  private closeRolesPopover(): void {
    this.clearRolesPopoverCloseTimer();
    if (!this.rolesPopoverEl) return;
    this.rolesPopoverEl.remove();
    this.rolesPopoverEl = null;
    this.rolesPopoverAnchor = null;
    document.removeEventListener('click', this.boundCloseRolesPopover);
    window.removeEventListener('scroll', this.boundCloseRolesPopover, true);
    window.removeEventListener('resize', this.boundCloseRolesPopover);
  }

  private toggleRolesPopover(anchor: HTMLElement, names: string[]): void {
    const reopeningSameAnchor = this.rolesPopoverAnchor === anchor;
    this.closeRolesPopover();
    if (reopeningSameAnchor) return;
    this.openRolesPopover(anchor, names);
  }

  private openRolesPopover(anchor: HTMLElement, names: string[]): void {
    this.closeRolesPopover();

    const popover = document.createElement('div');
    popover.className = 'roles-popover';
    popover.innerHTML = `
      <div class="roles-popover-header">Roles Assigned</div>
      <div class="roles-popover-list">
        ${names
          .map(
            (n) =>
              `<div class="roles-popover-item"><span class="roles-popover-dot"></span>${this.escapeHtml(n)}</div>`
          )
          .join('')}
      </div>
    `;

    // Hovering the popover itself counts as still hovering the badge —
    // cancels the close timer started by leaving the badge. Leaving the
    // popover schedules the same delayed close as leaving the badge.
    popover.addEventListener('mouseenter', () => this.clearRolesPopoverCloseTimer());
    popover.addEventListener('mouseleave', () => this.scheduleRolesPopoverClose());

    document.body.appendChild(popover);

    // Position (and re-clamp to viewport) only after it's in the DOM,
    // since that's the first point its real size is known.
    const anchorRect = anchor.getBoundingClientRect();
    const popRect = popover.getBoundingClientRect();

    let left = anchorRect.left;
    if (left + popRect.width > window.innerWidth - 8) {
      left = Math.max(8, window.innerWidth - popRect.width - 8);
    }
    let top = anchorRect.bottom + 6;
    if (top + popRect.height > window.innerHeight - 8) {
      top = anchorRect.top - popRect.height - 6;
    }
    popover.style.left = `${left}px`;
    popover.style.top = `${top}px`;

    this.rolesPopoverEl = popover;
    this.rolesPopoverAnchor = anchor;

    // Deferred so the click/hover that opened the popover doesn't
    // immediately bubble up and close it via this same listener.
    setTimeout(() => {
      document.addEventListener('click', this.boundCloseRolesPopover);
      window.addEventListener('scroll', this.boundCloseRolesPopover, true);
      window.addEventListener('resize', this.boundCloseRolesPopover);
    }, 0);
  }

  private buildActionColumn(): ColDef {
    return {
      colId: this.ACTION_COL_ID,
      field: this.ACTION_COL_ID,
      headerName: 'Action',
      sortable: false,
      filter: false,
      resizable: false,
      pinned: 'right',
      minWidth: 110,
      cellClass: 'action-cell-grid',
      cellRenderer: () => `
        <button type="button" class="btn-modify-grid">
          <img src="Images/modify.png" alt="Modify" width="20" height="20">
          Modify
        </button>`
    };
  }

  onGridCellClicked(event: any): void {
    if (event.colDef?.colId === this.ACTION_COL_ID) {
      this.onModify(event.data);
    }
  }

  onPaginationChanged(): void {
    if (!this.gridApi) return;
    const newPageSize = this.gridApi.paginationGetPageSize();
    if (!this.isGridReady) return;
    if (newPageSize !== this.PAGE_SIZE) {
      this.PAGE_SIZE = newPageSize;
      this.setDatasource();
    }
  }

  // ── Roles side panel (unchanged — separate, non-paginated /roles endpoint) ─
  loadRoles(): void {
    this.loadingRoles = true;
    this.service.getRoles().subscribe({
      next: (res) => {
        this.roles = (res?.roles ?? []).map(normalizeRole);
        this.loadingRoles = false;
      },
      error: (err) => {
        this.toastr.error(this.extractError(err), 'Error');
        this.loadingRoles = false;
      },
    });
  }

  private extractError(err: any): string {
    return err?.error?.detail || 'Something went wrong talking to the server.';
  }

  // ----- Modal state -----
  modalMode: ModalMode = null;
  editingUser: { id: number } | null = null;
  // Role IDs the user being edited had BEFORE this modal session — needed
  // to diff against configForm.roleIds on submit. The grid row itself
  // only carries role NAMES (backend stringifies roles for filtering —
  // see /users/list), so this is resolved once, at onModify() time, by
  // matching those names against the separately-loaded `this.roles` list
  // (which has real ids). Assumes role names are unique.
  private editingUserOriginalRoleIds: number[] = [];

  configForm = {
    userName: '',
    name: '',
    active: true,
    roleIds: [] as number[],
  };

  roleForm = {
    roleName: '',
  };

  roleToDelete: AdminRole | null = null;
  deletingRole = false;

  get isModalOpen(): boolean {
    return this.modalMode !== null;
  }

  isRoleChecked(roleId: number): boolean {
    return this.configForm.roleIds.includes(roleId);
  }

  toggleConfigRole(roleId: number, checked: boolean): void {
    if (checked) {
      if (!this.configForm.roleIds.includes(roleId)) {
        this.configForm.roleIds = [...this.configForm.roleIds, roleId];
      }
    } else {
      this.configForm.roleIds = this.configForm.roleIds.filter((id) => id !== roleId);
    }
  }

  // ----- Open modal: Add Configuration -----
  onAddConfiguration(): void {
    this.editingUser = null;
    this.editingUserOriginalRoleIds = [];
    this.configForm = {
      userName: '',
      name: '',
      active: true,
      roleIds: [],
    };
    this.modalMode = 'configuration';
  }

  // ----- Open modal: Modify Configuration (row is a raw ag-grid row) -----
  onModify(row: any): void {
    const rolesKey = Object.keys(row ?? {}).find((k) =>
      k.toLowerCase().replace(/[_\s]/g, '').includes('role')
    );
    const roleNames: string[] = this.rolesToNames(rolesKey ? row[rolesKey] : row.roles);

    const roleIds = this.roles
      .filter((r) => roleNames.includes(r.role_name))
      .map((r) => r.id);

    this.editingUser = { id: row.id };
    this.editingUserOriginalRoleIds = roleIds;
    this.configForm = {
      userName: row.username ?? row.userName ?? '',
      name: row.name ?? '',
      active: this.resolveActiveState(row),
      roleIds: [...roleIds],
    };
    this.modalMode = 'configuration';
  }

  /** Same "which field name did the backend actually use" uncertainty as
   *  buildColumnDefsFromMetadata's isActiveField check — handles the
   *  current shape (status: "Active"/"Inactive" string) plus the older
   *  is_active/active boolean shape, in case either ever comes back. */
  private resolveActiveState(row: any): boolean {
    const raw = row.status ?? row.is_active ?? row.active;
    if (typeof raw === 'string') {
      return ['active', 'true', '1', 'yes'].includes(raw.toLowerCase());
    }
    return raw === true || raw === 1;
  }

  // ----- Open modal: Add Role -----
  onAddRole(): void {
    this.roleForm = { roleName: '' };
    this.modalMode = 'role';
  }

  closeModal(): void {
    this.modalMode = null;
    this.editingUser = null;
    this.editingUserOriginalRoleIds = [];
    this.submitting = false;
  }

  // ----- Submit: Add/Modify Configuration -----
  private readonly EMAIL_PATTERN = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

  onSubmitConfig(): void {
    const userName = this.configForm.userName.trim();
    const name = this.configForm.name.trim();

    // Username validation
    if (!userName) {
      this.toastr.error('Email Address is required.', 'Validation Error');
      return;
    }

    if (!this.editingUser && !this.EMAIL_PATTERN.test(userName)) {
      this.toastr.error('Please enter a valid email address.', 'Validation Error');
      return;
    }

    // Name validation
    if (!name) {
      this.toastr.error('Full Name is required.', 'Validation Error');
      return;
    }

    this.submitting = true;

    if (this.editingUser) {
      const userId = this.editingUser.id;
      const previousRoleIds = this.editingUserOriginalRoleIds;
      const nextRoleIds = this.configForm.roleIds;
      const toAdd = nextRoleIds.filter((id) => !previousRoleIds.includes(id));
      const toRemove = previousRoleIds.filter((id) => !nextRoleIds.includes(id));

      this.service
        .updateUser(userId, { name: this.configForm.name.trim(), is_active: this.configForm.active })
        .subscribe({
          next: () => this.applyRoleDiff(userId, toAdd, toRemove),
          error: (err) => {
            this.toastr.error(this.extractError(err), 'Error');
            this.submitting = false;
          },
        });
    } else {
      this.service
        .createUser({
          username: this.configForm.userName.trim(),
          password: '',
          name: this.configForm.name.trim(),
          is_active: this.configForm.active,
          role_ids: this.configForm.roleIds,
        })
        .subscribe({
          next: () => {
            this.refreshGridData();
            this.closeModal();
            this.toastr.success('User Added successfully!', 'Success');
          },
          error: (err) => {
            this.toastr.error(this.extractError(err), 'Error');
            this.submitting = false;
          },
        });
    }
  }

  // Applies role assignment changes one at a time (the backend only
  // exposes single add/remove endpoints, no bulk-set endpoint).
  private applyRoleDiff(userId: number, toAdd: number[], toRemove: number[]): void {
    const calls: any[] = [
      ...toAdd.map((roleId) => this.service.assignUserRole(userId, { role_id: roleId })),
      ...toRemove.map((roleId) => this.service.unassignUserRole(userId, roleId)),
    ];

    if (calls.length === 0) {
      this.refreshGridData();
      this.closeModal();
      this.toastr.success('User updated successfully!', 'Success');
      return;
    }

    let remaining = calls.length;
    let hadError = false;
    calls.forEach((call) => {
      call.subscribe({
        next: () => {
          remaining -= 1;
          if (remaining === 0) {
            this.refreshGridData();
            if (!hadError) {
              this.closeModal();
              this.toastr.success('User updated successfully!', 'Success');
            } else {
              this.submitting = false;
            }
          }
        },
        error: (err: any) => {
          hadError = true;
          this.toastr.error(this.extractError(err), 'Error');
          remaining -= 1;
          if (remaining === 0) {
            this.refreshGridData();
            this.submitting = false;
          }
        },
      });
    });
  }

  // ----- Submit: Add Role -----
  onSubmitRole(): void {
    const roleName = this.roleForm.roleName.trim();
    if (!roleName) {
      this.toastr.error('Role name is required.', 'Validation Error');
      return;
    }

    this.submitting = true;
    this.service.createRole({ role_name: roleName }).subscribe({
      next: () => {
        this.loadRoles();
        this.closeModal();
        this.toastr.success('Role Added successfully!', 'Success');
      },
      error: (err) => {
        this.toastr.error(this.extractError(err), 'Error');
        this.submitting = false;
      },
    });
  }

  // ----- Delete role (with confirmation) -----
  onDeleteRoleClick(role: AdminRole): void {
    this.roleToDelete = role;
  }

  cancelDeleteRole(): void {
    this.roleToDelete = null;
    this.deletingRole = false;
  }

  confirmDeleteRole(): void {
    if (!this.roleToDelete) return;
    this.deletingRole = true;
    this.service.deleteRole(this.roleToDelete.id).subscribe({
      next: () => {
        this.roleToDelete = null;
        this.deletingRole = false;
        this.loadRoles();
        // A deleted role can change what shows in the Roles column/badges
        // for any user who held it, so refresh the grid too.
        this.refreshGridData();
        this.toastr.success('Role deleted successfully!', 'Success');
      },
      error: (err) => {
        this.toastr.error(this.extractError(err), 'Error');
        this.deletingRole = false;
      },
    });
  }
}