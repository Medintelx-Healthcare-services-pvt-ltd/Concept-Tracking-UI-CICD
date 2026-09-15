import { CommonModule } from '@angular/common';
import { HttpErrorResponse } from '@angular/common/http';
import { Component, effect, signal } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { LayoutComponent } from '../../layout/layout/layout';
import { ToastrService } from 'ngx-toastr';
import { Service } from '../../dashboard/service';

type CategoryType = 'coded' | 'plain' | 'transitions' | 'permissions';

interface ExtraField {
  key: string;         // matches a column name (description, claim_other, example)
  label: string;
  placeholder?: string;
  // Max length for this column — NOT hardcoded here. Left undefined until
  // loadItems() populates it from the API response's extraColLengths for
  // this key, so it always reflects whatever the backend currently
  // enforces rather than a static guess that can drift out of sync.
  maxLength?: number;
}

interface ConfigCategory {
  key: string;
  label: string;
  type: CategoryType;
  // Must match resolve_coded()/resolve_plain() in master_data_queries.py.
  apiSlug: string;
  // Only used for coded categories — seeded here as a best-guess default
  // (matches CODED_CONFIG's code_length in master_data_queries.py as of
  // this writing) but ALWAYS overwritten by the API response's own
  // codeLength on every loadItems() call — see there. Kept here only so
  // the Code field has a sane maxlength/hint before the first load
  // resolves.
  codeLength?: number;
  codePlaceholder?: string;
  // Populated from the API response's nameMaxLength on every loadItems()
  // call — not hardcoded, since it's authoritative from the backend and
  // can differ per category (and can change server-side independently
  // of this file).
  nameMaxLength?: number;
  // Extra columns this category's table carries beyond code/name — must
  // match CODED_CONFIG's extra_cols for the same category. Empty/undefined
  // for review-type and claim-type, which have no columns left over.
  // Each field's own maxLength is populated from the API response's
  // extraColLengths, same reasoning as nameMaxLength above.
  extraFields?: ExtraField[];
}

interface MasterDataItem {
  name: string;
  code?: string;        // present on coded items
  id?: number;           // present on plain items
  is_active?: number;
  description?: string;  // clients, master_concepts
  claim_other?: string;  // clients only
  example?: string;      // master_concepts only
}

interface RoleOption {
  role_id: number;
  role_name: string;
}

// ── Role Permissions matrix rows ─────────────────────────────────────────
// Shape matches GET /api/user-management/role-permissions's response —
// see role_permission_queries.py::fetch_*_matrix_for_role() on the backend.
// `isOverridden` distinguishes "this role has an explicit row" from
// "this is just the fallback default nobody's configured" — drives the
// Default/Custom badge and whether the Reset button shows.
interface ActionPermRow {
  permissionKey: string;
  isAllowed: boolean;
  isOverridden: boolean;
  default: boolean;
}

interface FieldPermRow {
  fieldName: string;
  accessLevel: 'edit' | 'view';
  isOverridden: boolean;
  default: 'edit' | 'view';
}

interface AttachmentPermRow {
  category: string;
  canManage: boolean;
  isOverridden: boolean;
  default: boolean;
}

@Component({
  selector: 'app-masterdataconfig',
  imports: [CommonModule, FormsModule, LayoutComponent],
  templateUrl: './masterdataconfig.html',
  styleUrl: './masterdataconfig.css',
})
export class Masterdataconfig {
  categories: ConfigCategory[] = [
    {
      key: 'clientName',
      label: 'Client Name',
      type: 'coded',
      apiSlug: 'client-name',
      codeLength: 3,
      codePlaceholder: 'e.g. CSP',
      extraFields: [
        { key: 'description', label: 'Description' },
        { key: 'claim_other', label: 'Claim Other' },
      ],
    },
    {
      key: 'masterConceptName',
      label: 'Master Concept Name',
      type: 'coded',
      apiSlug: 'master-concept-name',
      codeLength: 4,
      codePlaceholder: 'e.g. 0000',
      extraFields: [
        { key: 'description', label: 'Description' },
        { key: 'example', label: 'Example' },
      ],
    },
    { key: 'reviewType', label: 'Review Type', type: 'coded', apiSlug: 'review-type', codeLength: 1, codePlaceholder: 'e.g. A' },
    { key: 'claimType', label: 'Claim Type', type: 'coded', apiSlug: 'claim-type', codeLength: 1, codePlaceholder: 'e.g. P' },
    { key: 'priority', label: 'Priority', type: 'plain', apiSlug: 'priority' },
    { key: 'clientApprovalStatus', label: 'Client Approval Status', type: 'plain', apiSlug: 'client-approval-status' },
    { key: 'developmentStatus', label: 'Development Status', type: 'plain', apiSlug: 'development-status' },
    { key: 'statusTransitions', label: 'Status Transitions', type: 'transitions', apiSlug: '' },
    // { key: 'rolePermissions', label: 'Role Permissions', type: 'permissions', apiSlug: '' },
  ];

  activeCategory = signal<string>(this.categories[0].key);

  items = signal<MasterDataItem[]>([]);
  isLoading = signal(false);

  // ---- Add Configuration modal state ----
  isAddModalOpen = signal(false);
  newItemName = '';
  newItemCode = ''; // only used for coded categories
  newItemExtra: Record<string, string> = {}; // keyed by ExtraField.key
  isSaving = signal(false);

  // ---- Remove/deactivate confirm state ----
  pendingRemoveItem = signal<MasterDataItem | null>(null);
  isRemoving = signal(false);

  // ---- Status Transitions state ----
  transitionRoles = signal<RoleOption[]>([]);
  transitionStatuses = signal<string[]>([]);
  selectedRoleId = signal<number | null>(null);
  // "from||to" -> transition id, for existing edges of the selected role
  transitionEdgeMap = signal<Map<string, number>>(new Map());
  canOverride = signal(false);
  isLoadingTransitions = signal(false);
  isTogglingOverride = signal(false);
  // Non-null while the "turn override ON" confirm modal is open. Only used
  // for enabling — disabling override is always safe (it only restricts
  // further) so that path skips confirmation and calls onOverrideToggle
  // directly.
  pendingOverrideChange = signal<boolean | null>(null);
  // "from||to" keys currently mid-request, to disable that one checkbox
  pendingEdgeKeys = signal<Set<string>>(new Set());

  // ---- Role Permissions state ----
  permissionRoles = signal<RoleOption[]>([]);
  selectedPermRoleId = signal<number | null>(null);
  actionPerms = signal<ActionPermRow[]>([]);
  fieldPerms = signal<FieldPermRow[]>([]);
  attachmentPerms = signal<AttachmentPermRow[]>([]);
  isLoadingPermissions = signal(false);
  // "actions::key" / "fields::name" / "attachments::category" currently
  // mid-request, to disable that one control while its call is in flight.
  pendingPermKeys = signal<Set<string>>(new Set());

  constructor(
    private service: Service,
    private toastr: ToastrService,
  ) {
    // Reload the list any time the active category changes.
    effect(() => {
      const key = this.activeCategory();
      const category = this.categories.find((c) => c.key === key);
      if (category?.type === 'transitions') {
        this.initTransitionsView();
      } else if (category?.type === 'permissions') {
        this.initPermissionsView();
      } else {
        this.loadItems(key);
      }
    });
  }

  get activeCategoryMeta(): ConfigCategory {
    return this.categories.find((c) => c.key === this.activeCategory())!;
  }

  get activeCategoryLabel(): string {
    return this.activeCategoryMeta?.label ?? '';
  }

  get activeItems(): MasterDataItem[] {
    return this.items();
  }

  itemKey(_index: number, item: MasterDataItem): string | number {
    return item.code ?? item.id ?? item.name;
  }

  // Every coded category's Code field is alpha-only EXCEPT
  // master-concept-name, which is numeric-only (4-digit codes like
  // "0000") — every other coded category (client-name, review-type,
  // claim-type) uses letters only, per the sample data (MRW, CSP, A, P,
  // etc). Centralized here so both the live input-stripping handler and
  // the saveNewItem() validation check stay in sync — if this list of
  // "numeric" categories ever grows, only this needs to change.
  private isNumericCodeCategory(category: ConfigCategory): boolean {
    return category.key === 'masterConceptName';
  }

  /** Strips disallowed characters from the Code field as the user types —
   *  digits only for master-concept-name, letters only for every other
   *  coded category. Mirrors the pattern used elsewhere in this app for
   *  live input restriction (e.g. restrictSpecialChars in
   *  concept-create.ts). Doesn't block paste-then-submit on its own —
   *  saveNewItem()'s own check below is the real enforcement point. */
  onCodeInput(event: Event): void {
    const input = event.target as HTMLInputElement;
    const numeric = this.isNumericCodeCategory(this.activeCategoryMeta);
    const cleaned = numeric
      ? input.value.replace(/[^0-9]/g, '')
      : input.value.replace(/[^a-zA-Z]/g, '');
    input.value = cleaned;
    this.newItemCode = cleaned;
  }

  selectCategory(key: string): void {
    this.activeCategory.set(key);
  }

  private loadItems(key: string): void {
    const category = this.categories.find((c) => c.key === key);
    if (!category || category.type === 'transitions' || category.type === 'permissions') return;

    this.isLoading.set(true);

    const request$ =
      category.type === 'coded'
        ? this.service.getCodedMasterData(category.apiSlug)
        : this.service.getPlainMasterData(category.apiSlug);

    request$.subscribe({
      next: (res) => {
        this.items.set(res.items ?? []);

        // The backend is authoritative on these constraints — every
        // response carries its own current codeLength/nameMaxLength/
        // extraColLengths, so pull them in here rather than trusting the
        // static defaults seeded in `categories` above. Without this, a
        // frontend value that's drifted out of sync with the backend
        // (or a backend value that's changed since this file was last
        // updated) would keep silently enforcing the wrong length right
        // up until the create call 400s.
        if (category.type === 'coded' && res.codeLength) {
          category.codeLength = res.codeLength;
        }
        if (res.nameMaxLength) {
          category.nameMaxLength = res.nameMaxLength;
        }
        if (res.extraColLengths) {
          for (const field of category.extraFields ?? []) {
            const len = res.extraColLengths[field.key];
            if (len) field.maxLength = len;
          }
        }

        this.isLoading.set(false);
      },
      error: (err: HttpErrorResponse) => {
        this.items.set([]);
        this.isLoading.set(false);
        this.toastr.error(err.error?.detail ?? 'Failed to load entries.', 'Error');
      },
    });
  }

  // ---- Add ----

  openAddModal(): void {
    this.newItemName = '';
    this.newItemCode = '';
    this.newItemExtra = {};
    for (const field of this.activeCategoryMeta.extraFields ?? []) {
      this.newItemExtra[field.key] = '';
    }
    this.isAddModalOpen.set(true);
  }

  closeAddModal(): void {
    this.isAddModalOpen.set(false);
  }

  saveNewItem(): void {
    const category = this.activeCategoryMeta;
    const name = this.newItemName.trim();

    if (!name) {
      this.toastr.error('Name is required.', 'Validation Error');
      return;
    }

    // [maxlength] on the Name input blocks further typing once the limit
    // is hit, but doesn't stop a paste that lands over the limit in one
    // go — check explicitly here too, same as the codeLength check below.
    if (category.nameMaxLength && name.length > category.nameMaxLength) {
      this.toastr.error(
        `${category.label} name must be ${category.nameMaxLength} characters or fewer.`,
        'Validation Error',
      );
      return;
    }

    this.isSaving.set(true);

    if (category.type === 'coded') {
      const code = this.newItemCode.trim();
      if (!code) {
        this.isSaving.set(false);
        this.toastr.error('Code is required.', 'Validation Error');
        return;
      }
      if (category.codeLength && code.length !== category.codeLength) {
        this.isSaving.set(false);
        this.toastr.error(
          `${category.label} code must be exactly ${category.codeLength} character${category.codeLength === 1 ? '' : 's'}.`,
          'Validation Error',
        );
        return;
      }

      const numeric = this.isNumericCodeCategory(category);
      const codePattern = numeric ? /^[0-9]+$/ : /^[a-zA-Z]+$/;
      if (!codePattern.test(code)) {
        this.isSaving.set(false);
        this.toastr.error(
          numeric
            ? `${category.label} code must contain numbers only.`
            : `${category.label} code must contain letters only, no numbers or special characters.`,
          'Validation Error',
        );
        return;
      }

      const extraPayload: Record<string, string> = {};
      for (const field of category.extraFields ?? []) {
        const value = (this.newItemExtra[field.key] ?? '').trim();
        if (field.maxLength && value.length > field.maxLength) {
          this.isSaving.set(false);
          this.toastr.error(
            `${field.label} must be ${field.maxLength} characters or fewer.`,
            'Validation Error',
          );
          return;
        }
        extraPayload[field.key] = value;
      }

      this.service
        .createCodedMasterData(category.apiSlug, { code, name, is_active: 1, ...extraPayload })
        .subscribe({
          next: () => {
            this.isSaving.set(false);
            this.closeAddModal();
            this.loadItems(category.key);
            this.toastr.success(`${category.label} added successfully!`, 'Success');
          },
          error: (err: HttpErrorResponse) => {
            this.isSaving.set(false);
            this.toastr.error(err.error?.detail ?? 'Failed to save.', 'Error');
          },
        });
    } else {
      this.service.createPlainMasterData(category.apiSlug, { name }).subscribe({
        next: () => {
          this.isSaving.set(false);
          this.closeAddModal();
          this.loadItems(category.key);
          this.toastr.success(`${category.label} added successfully!`, 'Success');
        },
        error: (err: HttpErrorResponse) => {
          this.isSaving.set(false);
          this.toastr.error(err.error?.detail ?? 'Failed to save.', 'Error');
        },
      });
    }
  }

  // ---- Remove / deactivate ----

  requestRemoveItem(item: MasterDataItem): void {
    this.pendingRemoveItem.set(item);
  }

  cancelRemoveItem(): void {
    this.pendingRemoveItem.set(null);
  }

  confirmRemoveItem(): void {
    const category = this.activeCategoryMeta;
    const item = this.pendingRemoveItem();
    if (!item) return;

    if (category.type === 'coded' && !item.code) {
      this.toastr.success(`${category.label} deleted successfully!`, 'Success');
      return;
    }
    if (category.type === 'plain' && item.id == null) {
      this.toastr.success(`${category.label} deleted successfully!`, 'Success');
      return;
    }

    this.isRemoving.set(true);

    const request$ =
      category.type === 'coded'
        ? this.service.deleteCodedMasterData(category.apiSlug, item.code!)
        : this.service.deletePlainMasterData(category.apiSlug, item.id!);

    request$.subscribe({
      next: () => {
        this.isRemoving.set(false);
        this.pendingRemoveItem.set(null);
        this.loadItems(category.key);
        this.toastr.success('value deleted successfully!', 'Success');
      },
      error: (err: HttpErrorResponse) => {
        this.isRemoving.set(false);
        this.toastr.error(err.error?.detail ?? 'Failed to delete.', 'Error');
      }
    });
  }

  // =====================================================================
  // STATUS TRANSITIONS
  // =====================================================================

  private edgeKey(from: string, to: string): string {
    return `${from}||${to}`;
  }

  /** Loads roles + statuses once (both needed to draw the matrix), then
   *  loads transitions for whichever role is currently selected — or
   *  defaults to the first role if none was picked yet. */
  private initTransitionsView(): void {
    this.isLoadingTransitions.set(true);

    this.service.getRoles().subscribe({
      next: (res) => {
        const roles: RoleOption[] = res.roles ?? [];
        this.transitionRoles.set(roles);

        if (this.selectedRoleId() == null && roles.length > 0) {
          this.selectedRoleId.set(roles[0].role_id);
        }

        this.service.getPlainMasterData('development-status').subscribe({
          next: (statusRes) => {
            const statuses: string[] = (statusRes.items ?? []).map((i: MasterDataItem) => i.name);
            this.transitionStatuses.set(statuses);

            const roleId = this.selectedRoleId();
            if (roleId != null) {
              this.loadTransitionsForRole(roleId);
            } else {
              this.isLoadingTransitions.set(false);
            }
          },
          error: (err: HttpErrorResponse) => {
            this.isLoadingTransitions.set(false);
            this.toastr.error(err.error?.detail ?? 'Failed to load statuses.', 'Error');
          },
        });
      },
      error: (err: HttpErrorResponse) => {
        this.isLoadingTransitions.set(false);
        this.toastr.error(err.error?.detail ?? 'Failed to load roles.', 'Error');
      },
    });
  }

  onRoleSelectChange(roleIdStr: string): void {
    const roleId = Number(roleIdStr);
    this.selectedRoleId.set(roleId);
    this.loadTransitionsForRole(roleId);
  }

  private loadTransitionsForRole(roleId: number): void {
    this.isLoadingTransitions.set(true);
    this.service.getStatusTransitions(roleId).subscribe({
      next: (res) => {
        const map = new Map<string, number>();
        for (const t of res.transitions ?? []) {
          map.set(this.edgeKey(t.from_status, t.to_status), t.id);
        }
        this.transitionEdgeMap.set(map);
        this.canOverride.set(!!res.canOverride);
        this.isLoadingTransitions.set(false);
      },
      error: (err: HttpErrorResponse) => {
        this.isLoadingTransitions.set(false);
        this.toastr.error(err.error?.detail ?? 'Failed to load transitions.', 'Error');
      },
    });
  }

  isEdgeChecked(from: string, to: string): boolean {
    return this.transitionEdgeMap().has(this.edgeKey(from, to));
  }

  isEdgePending(from: string, to: string): boolean {
    return this.pendingEdgeKeys().has(this.edgeKey(from, to));
  }

  toggleEdge(from: string, to: string, checked: boolean): void {
    const roleId = this.selectedRoleId();
    if (roleId == null || from === to) return;

    const key = this.edgeKey(from, to);
    const pending = new Set(this.pendingEdgeKeys());
    pending.add(key);
    this.pendingEdgeKeys.set(pending);

    const done = () => {
      const p = new Set(this.pendingEdgeKeys());
      p.delete(key);
      this.pendingEdgeKeys.set(p);
    };

    if (checked) {
      this.service.addStatusTransition({ role_id: roleId, from_status: from, to_status: to }).subscribe({
        next: (res) => {
          const map = new Map(this.transitionEdgeMap());
          map.set(key, res.id);
          this.transitionEdgeMap.set(map);
          done();
        },
        error: (err: HttpErrorResponse) => {
          done();
          this.toastr.error(err.error?.detail ?? 'Failed to add transition.', 'Error');
        },
      });
    } else {
      const transitionId = this.transitionEdgeMap().get(key);
      if (transitionId == null) {
        done();
        return;
      }
      this.service.removeStatusTransition(transitionId).subscribe({
        next: () => {
          const map = new Map(this.transitionEdgeMap());
          map.delete(key);
          this.transitionEdgeMap.set(map);
          done();
        },
        error: (err: HttpErrorResponse) => {
          done();
          this.toastr.error(err.error?.detail ?? 'Failed to remove transition.', 'Error');
        },
      });
    }
  }

  /** Bound to the checkbox's (change) event. Turning override ON is
   *  destructive to the role's transition restrictions, so it opens a
   *  confirm modal instead of calling the API immediately — the checkbox
   *  itself snaps back to canOverride()'s current value on the next change
   *  detection cycle since nothing here mutates that signal yet. Turning
   *  override OFF is safe (never grants access) so it applies right away. */
  onOverrideCheckboxChange(checked: boolean): void {
    if (checked) {
      this.pendingOverrideChange.set(true);
    } else {
      this.onOverrideToggle(false);
    }
  }

  get pendingOverrideRoleName(): string {
    const roleId = this.selectedRoleId();
    return this.transitionRoles().find((r) => r.role_id === roleId)?.role_name ?? 'this role';
  }

  cancelOverrideChange(): void {
    this.pendingOverrideChange.set(null);
  }

  confirmOverrideChange(): void {
    this.pendingOverrideChange.set(null);
    this.onOverrideToggle(true);
  }

  onOverrideToggle(checked: boolean): void {
    const roleId = this.selectedRoleId();
    if (roleId == null) return;

    this.isTogglingOverride.set(true);
    this.service.setRoleOverride(roleId, checked).subscribe({
      next: () => {
        this.canOverride.set(checked);
        this.isTogglingOverride.set(false);
        this.toastr.success(
          checked ? 'This role can now set any status.' : 'Override removed for this role.',
          'Success',
        );
      },
      error: (err: HttpErrorResponse) => {
        this.isTogglingOverride.set(false);
        this.toastr.error(err.error?.detail ?? 'Failed to update override.', 'Error');
      },
    });
  }

  // =====================================================================
  // ROLE PERMISSIONS
  // =====================================================================

  /** Loads roles once, then loads the full actions/fields/attachments
   *  matrix for whichever role is currently selected — or defaults to the
   *  first role if none was picked yet. Mirrors initTransitionsView(). */
  private initPermissionsView(): void {
    this.isLoadingPermissions.set(true);

    this.service.getRoles().subscribe({
      next: (res) => {
        const roles: RoleOption[] = res.roles ?? [];
        this.permissionRoles.set(roles);

        if (this.selectedPermRoleId() == null && roles.length > 0) {
          this.selectedPermRoleId.set(roles[0].role_id);
        }

        const roleId = this.selectedPermRoleId();
        if (roleId != null) {
          this.loadPermissionsForRole(roleId);
        } else {
          this.isLoadingPermissions.set(false);
        }
      },
      error: (err: HttpErrorResponse) => {
        this.isLoadingPermissions.set(false);
        this.toastr.error(err.error?.detail ?? 'Failed to load roles.', 'Error');
      },
    });
  }

  onPermRoleSelectChange(roleIdStr: string): void {
    const roleId = Number(roleIdStr);
    this.selectedPermRoleId.set(roleId);
    this.loadPermissionsForRole(roleId);
  }

  private loadPermissionsForRole(roleId: number): void {
    this.isLoadingPermissions.set(true);
    this.service.getRolePermissionMatrix(roleId).subscribe({
      next: (res) => {
        this.actionPerms.set(res.actions ?? []);
        this.fieldPerms.set(res.fields ?? []);
        this.attachmentPerms.set(res.attachments ?? []);
        this.isLoadingPermissions.set(false);
      },
      error: (err: HttpErrorResponse) => {
        this.isLoadingPermissions.set(false);
        this.toastr.error(err.error?.detail ?? 'Failed to load permissions.', 'Error');
      },
    });
  }

  private permPendingKey(kind: 'actions' | 'fields' | 'attachments', key: string): string {
    return `${kind}::${key}`;
  }

  isPermPending(kind: 'actions' | 'fields' | 'attachments', key: string): boolean {
    return this.pendingPermKeys().has(this.permPendingKey(kind, key));
  }

  private setPermPending(kind: 'actions' | 'fields' | 'attachments', key: string, pending: boolean): void {
    const set = new Set(this.pendingPermKeys());
    const k = this.permPendingKey(kind, key);
    if (pending) set.add(k);
    else set.delete(k);
    this.pendingPermKeys.set(set);
  }

  // ---- Actions ----

  toggleActionPermission(row: ActionPermRow, checked: boolean): void {
    const roleId = this.selectedPermRoleId();
    if (roleId == null) return;

    this.setPermPending('actions', row.permissionKey, true);
    this.service.setActionPermission(roleId, row.permissionKey, checked).subscribe({
      next: () => {
        this.actionPerms.set(
          this.actionPerms().map((r) =>
            r.permissionKey === row.permissionKey ? { ...r, isAllowed: checked, isOverridden: true } : r,
          ),
        );
        this.setPermPending('actions', row.permissionKey, false);
      },
      error: (err: HttpErrorResponse) => {
        this.setPermPending('actions', row.permissionKey, false);
        this.toastr.error(err.error?.detail ?? 'Failed to update permission.', 'Error');
      },
    });
  }

  resetActionPermission(row: ActionPermRow): void {
    const roleId = this.selectedPermRoleId();
    if (roleId == null) return;

    this.setPermPending('actions', row.permissionKey, true);
    this.service.resetActionPermission(roleId, row.permissionKey).subscribe({
      next: () => {
        this.actionPerms.set(
          this.actionPerms().map((r) =>
            r.permissionKey === row.permissionKey
              ? { ...r, isAllowed: row.default, isOverridden: false }
              : r,
          ),
        );
        this.setPermPending('actions', row.permissionKey, false);
        this.toastr.success('Reset to default.', 'Success');
      },
      error: (err: HttpErrorResponse) => {
        this.setPermPending('actions', row.permissionKey, false);
        this.toastr.error(err.error?.detail ?? 'Failed to reset permission.', 'Error');
      },
    });
  }

  // ---- Fields ----

  onFieldAccessChange(row: FieldPermRow, accessLevel: 'edit' | 'view'): void {
    const roleId = this.selectedPermRoleId();
    if (roleId == null) return;

    this.setPermPending('fields', row.fieldName, true);
    this.service.setFieldPermission(roleId, row.fieldName, accessLevel).subscribe({
      next: () => {
        this.fieldPerms.set(
          this.fieldPerms().map((r) =>
            r.fieldName === row.fieldName ? { ...r, accessLevel, isOverridden: true } : r,
          ),
        );
        this.setPermPending('fields', row.fieldName, false);
      },
      error: (err: HttpErrorResponse) => {
        this.setPermPending('fields', row.fieldName, false);
        this.toastr.error(err.error?.detail ?? 'Failed to update field permission.', 'Error');
      },
    });
  }

  resetFieldPermission(row: FieldPermRow): void {
    const roleId = this.selectedPermRoleId();
    if (roleId == null) return;

    this.setPermPending('fields', row.fieldName, true);
    this.service.resetFieldPermission(roleId, row.fieldName).subscribe({
      next: () => {
        this.fieldPerms.set(
          this.fieldPerms().map((r) =>
            r.fieldName === row.fieldName ? { ...r, accessLevel: row.default, isOverridden: false } : r,
          ),
        );
        this.setPermPending('fields', row.fieldName, false);
        this.toastr.success('Reset to default.', 'Success');
      },
      error: (err: HttpErrorResponse) => {
        this.setPermPending('fields', row.fieldName, false);
        this.toastr.error(err.error?.detail ?? 'Failed to reset field permission.', 'Error');
      },
    });
  }

  // ---- Attachments ----

  toggleAttachmentPermission(row: AttachmentPermRow, checked: boolean): void {
    const roleId = this.selectedPermRoleId();
    if (roleId == null) return;

    this.setPermPending('attachments', row.category, true);
    this.service.setAttachmentPermission(roleId, row.category, checked).subscribe({
      next: () => {
        this.attachmentPerms.set(
          this.attachmentPerms().map((r) =>
            r.category === row.category ? { ...r, canManage: checked, isOverridden: true } : r,
          ),
        );
        this.setPermPending('attachments', row.category, false);
      },
      error: (err: HttpErrorResponse) => {
        this.setPermPending('attachments', row.category, false);
        this.toastr.error(err.error?.detail ?? 'Failed to update attachment permission.', 'Error');
      },
    });
  }

  resetAttachmentPermission(row: AttachmentPermRow): void {
    const roleId = this.selectedPermRoleId();
    if (roleId == null) return;

    this.setPermPending('attachments', row.category, true);
    this.service.resetAttachmentPermission(roleId, row.category).subscribe({
      next: () => {
        this.attachmentPerms.set(
          this.attachmentPerms().map((r) =>
            r.category === row.category ? { ...r, canManage: row.default, isOverridden: false } : r,
          ),
        );
        this.setPermPending('attachments', row.category, false);
        this.toastr.success('Reset to default.', 'Success');
      },
      error: (err: HttpErrorResponse) => {
        this.setPermPending('attachments', row.category, false);
        this.toastr.error(err.error?.detail ?? 'Failed to reset attachment permission.', 'Error');
      },
    });
  }
}