import { Injectable } from '@angular/core';
import { HttpClient, HttpHeaders, HttpParams } from '@angular/common/http';
import {  Observable } from 'rxjs';
import { AppConfig } from '../app.config';
import { MyRolesResponse } from '../layout/layout/layout';

@Injectable({
  providedIn: 'root',
})
export class Service {
  
  constructor(private http: HttpClient) {
    this.apiBaseUrl = AppConfig.settings?.apiurl || '';
  }
  
  public apiBaseUrl: string;
  
  getPdfFile(url: string): Observable<Blob> {
    const fullUrl = `${this.apiBaseUrl}${url}`;
    return this.http.post(fullUrl, {}, {
      responseType: 'blob'
    });
  }
 

  clear() {
    sessionStorage.clear();
  }

  // ----------------

  getdashboardConcepts(requestBody: any): Observable<any> {
    const url = `${this.apiBaseUrl}api/dashboard-concepts`;
    return this.http.post<any>(url, requestBody);
  }

  getmasterdata(): Observable<any> {
    const url= `${this.apiBaseUrl}api/master-data`
    return this.http.get<any>(url);
  }

  /** Returns the statuses the CURRENT user's role is allowed to move a
   *  concept to, given its current status. Backed by the same
   *  status_permissions.py matrix the create-concept endpoint enforces —
   *  this is what the dropdown should be filtered against, not the raw
   *  developmentStatusOptions list from master-data. */
  getAllowedStatuses(userId: number, currentStatus: string, roleName: string): Observable<any> {
    const params = new HttpParams()
      .set('user_id', userId)
      .set('current_status', currentStatus || 'New')
      .set('role_name', roleName || '');
    const url = `${this.apiBaseUrl}api/allowed-statuses`;
    return this.http.get<any>(url, { params });
  }

  createconcept(formData: FormData): Observable<any> {
    const url= `${this.apiBaseUrl}api/create-concept`
    return this.http.post<any>(url,formData);
  }

  /** Second step of login for a user with more than one role - exchanges
   *  the short-lived pending_token + chosen role_id for a real session
   *  token via POST /select-role. */
  selectRole(pendingToken: string, roleId: number): Observable<any> {
    const url = `${this.apiBaseUrl}select-role`;
    return this.http.post<any>(url, { pending_token: pendingToken, role_id: roleId });
  }

  /** Mid-session role change - the current valid token proves identity,
   *  no password re-entry needed. Returns a new token under the new role. */
  switchRole(roleId: number): Observable<any> {
    const url = `${this.apiBaseUrl}switch-role`;
    return this.http.post<any>(url, { role_id: roleId });
  }

  getMyRoles(userId: number): Observable<MyRolesResponse> {
    const params = new HttpParams().set('user_id', userId);
    const url = `${this.apiBaseUrl}my-roles`;
    return this.http.get<MyRolesResponse>(url, { params });
  }
  // submitclientApproval(data: any): Observable<any> {
  //   return this.http.post(`${this.apiBaseUrl}api/submit_client_approval`, data);
  // }
  submitclientApproval(data: any, conceptId: string): Observable<any> {
  const url = `${this.apiBaseUrl}api/submit_client_approval/?${conceptId}`;
  return this.http.post<any>(url, data);
}

  submitsupportingdocuments(formData: FormData): Observable<any> {
    const url = `${this.apiBaseUrl}api/upload_supporting_docs`;
    return this.http.post<any>(url,formData);
  }

  /** Fetch a concept by ID */
  getConcept(conceptId: string): Observable<any> {
    const url = `${this.apiBaseUrl}api/concepts/${conceptId}`;
    return this.http.get<any>(url);
  }

  getConceptHistory(conceptId: string): Observable<any> {
    const url = `${this.apiBaseUrl}api/concepts/${conceptId}/history`;
    return this.http.get<any>(url);
  }

  getLatestUpdates(): Observable<any> {
    return this.http.get<any>(`${this.apiBaseUrl}api/latest-updates`);
  }

  getConceptsByUserId(user_id: number): Observable<any> {
    const url = `${this.apiBaseUrl}api/Concept/${user_id}`;
    return this.http.get<any>(url);
  }

  /** Save / update concept core details */
  saveConcept(payload: any): Observable<any> {
    const headers = new HttpHeaders({ 'Content-Type': 'application/json' });
    return this.http.post(`${this.apiBaseUrl}api/concepts`, payload, { headers });
  }

  getuserfiles(user_id: number): Observable<any> {
    const url = `${this.apiBaseUrl}api/user-files/${user_id}`;
    return this.http.get<any>(url);
  }
  
  deleteattachment(formData: FormData): Observable<any> {
    const url = `${this.apiBaseUrl}api/delete-attachment`;
    return this.http.post<any>(url, formData);
  }

  // =====================================================================
  // MASTER DATA MANAGEMENT (admin config screen)
  // Backs master_data_management_route, prefix
  // /api/user-management/master-data. Coded categories (Client Name,
  // Master Concept Name, Review Type, Claim Type) carry a code + is_active
  // flag; plain categories (Development Status, Priority, Client Approval
  // Status) are name-only, no is_active, no delete route on the backend.
  // `category` must match whatever resolve_coded()/resolve_plain() expect
  // in master_data_queries.py.
  // =====================================================================

  getCodedMasterData(category: string): Observable<any> {
    const url = `${this.apiBaseUrl}api/user-management/master-data/coded/${category}`;
    return this.http.get<any>(url);
  }

  createCodedMasterData(category: string, payload: { code: string; name: string; is_active: number; [extra: string]: string | number }): Observable<any> {
    const url = `${this.apiBaseUrl}api/user-management/master-data/coded/${category}`;
    return this.http.post<any>(url, payload);
  }

  updateCodedMasterData(category: string, code: string, payload: { name: string; is_active: number; [extra: string]: string | number }): Observable<any> {
    const url = `${this.apiBaseUrl}api/user-management/master-data/coded/${category}/${code}`;
    return this.http.put<any>(url, payload);
  }

  deleteCodedMasterData(category: string, code: string): Observable<any> {
    const url = `${this.apiBaseUrl}api/user-management/master-data/coded/${category}/${code}`;
    return this.http.delete<any>(url);
  }

  getPlainMasterData(category: string): Observable<any> {
    const url = `${this.apiBaseUrl}api/user-management/master-data/plain/${category}`;
    return this.http.get<any>(url);
  }

  createPlainMasterData(category: string, payload: { name: string }): Observable<any> {
    const url = `${this.apiBaseUrl}api/user-management/master-data/plain/${category}`;
    return this.http.post<any>(url, payload);
  }

  updatePlainMasterData(category: string, itemId: number, payload: { name: string }): Observable<any> {
    const url = `${this.apiBaseUrl}api/user-management/master-data/plain/${category}/${itemId}`;
    return this.http.put<any>(url, payload);
  }

  deletePlainMasterData(category: string, itemId: number): Observable<any> {
    const url = `${this.apiBaseUrl}api/user-management/master-data/plain/${category}/${itemId}`;
    return this.http.delete<any>(url);
  }

  // =====================================================================
  // USER MANAGEMENT (admin: users, roles, role assignment)
  // Backs user_management_route, prefix /api/user-management.
  // Every route requires the caller to already hold the 'Admin' role.
  // =====================================================================

  getUsers(body: any): Observable<any> {
    const url = `${this.apiBaseUrl}api/user-management/users/list`;
    return this.http.post<any>(url,body);
  }

  createUser(payload: { username: string; password: string; name: string; is_active: boolean; role_ids: number[] }): Observable<any> {
    const url = `${this.apiBaseUrl}api/user-management/users`;
    return this.http.post<any>(url, payload);
  }

  updateUser(userId: number, payload: { name: string; is_active: boolean }): Observable<any> {
    const url = `${this.apiBaseUrl}api/user-management/users/${userId}`;
    return this.http.put<any>(url, payload);
  }

  getRoles(): Observable<any> {
    const url = `${this.apiBaseUrl}api/user-management/roles`;
    return this.http.get<any>(url);
  }

  createRole(payload: { role_name: string }): Observable<any> {
    const url = `${this.apiBaseUrl}api/user-management/roles`;
    return this.http.post<any>(url, payload);
  }

  // NOTE: this route doesn't exist yet on user_management_route (only
  // GET/POST /roles are defined there) — add a DELETE /roles/{role_id}
  // handler on the backend for this to work.
  deleteRole(roleId: number): Observable<any> {
    const url = `${this.apiBaseUrl}api/user-management/roles/${roleId}`;
    return this.http.delete<any>(url);
  }

  getUserRoles(userId: number): Observable<any> {
    const url = `${this.apiBaseUrl}api/user-management/users/${userId}/roles`;
    return this.http.get<any>(url);
  }

  assignUserRole(userId: number, payload: { role_id: number }): Observable<any> {
    const url = `${this.apiBaseUrl}api/user-management/users/${userId}/roles`;
    return this.http.post<any>(url, payload);
  }

  unassignUserRole(userId: number, roleId: number): Observable<any> {
    const url = `${this.apiBaseUrl}api/user-management/users/${userId}/roles/${roleId}`;
    return this.http.delete<any>(url);
  }

  copyConcept(conceptId: string) {
    return this.http.get(`${this.apiBaseUrl}api/copy-concept/${conceptId}`, {});
  }

  getStatusTransitions(roleId: number): Observable<any> {
    const url = `${this.apiBaseUrl}api/user-management/status-transitions`;
    return this.http.get<any>(url, { params: { role_id: roleId } });
  }

  addStatusTransition(payload: { role_id: number; from_status: string; to_status: string }): Observable<any> {
    const url = `${this.apiBaseUrl}api/user-management/status-transitions`;
    return this.http.post<any>(url, payload);
  }

  removeStatusTransition(transitionId: number): Observable<any> {
    const url = `${this.apiBaseUrl}api/user-management/status-transitions/${transitionId}`;
    return this.http.delete<any>(url);
  }

  setRoleOverride(roleId: number, canOverride: boolean): Observable<any> {
    const url = `${this.apiBaseUrl}api/user-management/status-transitions/roles/${roleId}/override`;
    return this.http.put<any>(url, { can_override_status: canOverride });
  }

  // =====================================================================
  // ROLE PERMISSIONS MANAGEMENT (admin config screen)
  // Backs role_permissions_admin_route, prefix
  // /api/user-management/role-permissions. Every route requires the
  // caller to already hold the 'Admin' role, same as the status-
  // transitions and master-data management blocks above.
  //
  // NOTE: this is separate from getRolePermissions() further up this
  // file, which hits the READ-ONLY, non-admin-gated resolver
  // (GET /api/role-permissions/{role_id}) that every logged-in user
  // calls for their own session to decide what to render.
  // =====================================================================
  /** Schema inventory for the admin grid — every known action key, field
   *  name, and attachment category. Same for every role; tells the UI
   *  what rows/columns to draw without hardcoding the list a second time
   *  in the frontend. */
  getPermissionKeys(): Observable<any> {
    const url = `${this.apiBaseUrl}api/user-management/role-permissions/keys`;
    return this.http.get<any>(url);
  }
  /** Full editable matrix for one role — every action/field/category,
   *  each with its current value and an `isOverridden` flag (explicit row
   *  vs. falling back to the schema default). This is what the admin
   *  screen renders as its grid. */
  getRolePermissionMatrix(roleId: number): Observable<any> {
    const url = `${this.apiBaseUrl}api/user-management/role-permissions`;
    return this.http.get<any>(url, { params: { role_id: roleId } });
  }
  setActionPermission(roleId: number, permissionKey: string, isAllowed: boolean): Observable<any> {
    const url = `${this.apiBaseUrl}api/user-management/role-permissions/${roleId}/actions/${permissionKey}`;
    return this.http.put<any>(url, { is_allowed: isAllowed });
  }
  /** Removes this role's explicit override for the key — "Reset to
   *  default" in the UI. Falls back to the schema's restrictive default,
   *  never to an unrestricted state. */
  resetActionPermission(roleId: number, permissionKey: string): Observable<any> {
    const url = `${this.apiBaseUrl}api/user-management/role-permissions/${roleId}/actions/${permissionKey}`;
    return this.http.delete<any>(url);
  }
  setFieldPermission(roleId: number, fieldName: string, accessLevel: 'edit' | 'view'): Observable<any> {
    const url = `${this.apiBaseUrl}api/user-management/role-permissions/${roleId}/fields/${fieldName}`;
    return this.http.put<any>(url, { access_level: accessLevel });
  }
  resetFieldPermission(roleId: number, fieldName: string): Observable<any> {
    const url = `${this.apiBaseUrl}api/user-management/role-permissions/${roleId}/fields/${fieldName}`;
    return this.http.delete<any>(url);
  }
  setAttachmentPermission(roleId: number, category: string, canManage: boolean): Observable<any> {
    const url = `${this.apiBaseUrl}api/user-management/role-permissions/${roleId}/attachments/${category}`;
    return this.http.put<any>(url, { can_manage: canManage });
  }
  resetAttachmentPermission(roleId: number, category: string): Observable<any> {
    const url = `${this.apiBaseUrl}api/user-management/role-permissions/${roleId}/attachments/${category}`;
    return this.http.delete<any>(url);
  }
}

