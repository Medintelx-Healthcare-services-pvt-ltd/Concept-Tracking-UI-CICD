import { Component, ElementRef, EventEmitter, HostListener, Input, OnInit, Output, ViewChild, inject } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { RouterModule, Router } from '@angular/router';
import { VERSION } from '../../../environments/version';
import { Service } from '../../dashboard/service'; 
import { ToastrService } from 'ngx-toastr';

export interface RoleOption {
  role_id: number;
  role_name: string;
}

export interface MyRolesResponse {
  roles: RoleOption[];
}
@Component({
  selector: 'app-layout',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    RouterModule
  ],
  templateUrl: './layout.html',
  styleUrls: ['./layout.css']
})
export class LayoutComponent implements OnInit {
  VERSION: string = VERSION.version;
  private router = inject(Router);
  private service = inject(Service);
  private toastr = inject(ToastrService);

  // Scoped specifically to the role-switcher button/menu — NOT the whole
  // app-layout host. The host wraps the entire page (sidebar, top nav,
  // and all page content via <ng-content>), so checking against
  // elementRef.nativeElement there meant almost any click anywhere in
  // the app still counted as "inside", and the dropdown never actually
  // saw a genuine outside click. This ref only covers the switcher's own
  // button + menu, set on the div in layout.html via #roleSwitcher.
  @ViewChild('roleSwitcher') roleSwitcherRef?: ElementRef<HTMLElement>;
  searchQuery: string = '';
  userInitials: string = '';

  // Left-panel collapse toggle — hidden by default, opted into by whichever
  // page (e.g. Concept Create) actually renders a collapsible left panel.
  @Input() showPanelToggle: boolean = false;
  @Input() panelCollapsed: boolean = false;
  @Output() panelCollapsedChange = new EventEmitter<boolean>();
  @Input() hideMainNav: boolean = false;
  @Input() showAdminLink: boolean = false;
  togglePanel(): void {
    this.panelCollapsed = !this.panelCollapsed;
    this.panelCollapsedChange.emit(this.panelCollapsed);
  }

  // ── Role switching ──────────────────────────────────────────────
  // availableRoles is only populated when the user actually has more
  // than one role - see the select-role page, which is the one place
  // the full role list is known and must cache it here for later use.
  availableRoles: { role_id: number; role_name: string }[] = [];
  currentRoleName: string = '';
  showRoleMenu = false;

  // Drives both the sidebar's Admin icon/disabled state (in layout.html)
  // and switchRole()'s redirect target below. Case-insensitive since the
  // backend's role_name casing ("Admin") shouldn't be load-bearing here.
  get isAdmin(): boolean {
    return (this.currentRoleName || '').toLowerCase() === 'admin';
  }

  // Reads the cached role list/current role out of sessionStorage only —
  // no network call. Used for the initial paint (before the dropdown has
  // ever been opened) so the nav bar doesn't flash empty while a request
  // is in flight.
  loadRoles(): void {
    try {
      const raw = sessionStorage.getItem('availableRoles');
      this.availableRoles = raw ? JSON.parse(raw) : [];
    } catch {
      this.availableRoles = [];
    }
    this.currentRoleName = sessionStorage.getItem('roleName') ?? '';
  }

  // Hits the backend for the current role list. Previously this ran on a
  // 1-minute setInterval regardless of whether anyone was even looking at
  // the switcher — polling the API for no reason most of the time. Now
  // it's only called when the dropdown is actually opened (see
  // toggleRoleMenu below), so the list is refreshed exactly when it's
  // about to be read/acted on, not on a fixed timer.
  private refreshRolesFromServer(): void {
    const userId = Number(sessionStorage.getItem('userId'));
    if (!userId) {
      return; // no valid session yet, nothing to refresh against
    }

    this.service.getMyRoles(userId).subscribe({
      next: (res: MyRolesResponse) => {
        this.availableRoles = res.roles ?? [];
        sessionStorage.setItem('availableRoles', JSON.stringify(this.availableRoles));
      },
      error: () => {
        // Keep the cached copy already set by loadRoles() as a fallback.
      },
    });
  }

  // Opening the dropdown is the one moment a stale role list actually
  // matters (about to pick from it), so that's the trigger for the
  // network refresh — closing it doesn't need one.
  toggleRoleMenu(): void {
    this.showRoleMenu = !this.showRoleMenu;
    if (this.showRoleMenu) {
      this.refreshRolesFromServer();
    }
  }

  // Close the menu the moment an option is picked, rather than waiting on
  // the switchRole HTTP call to resolve — otherwise the dropdown lingers
  // open (visibly, for however long the request takes) before the reload
  // below finally kicks in and wipes it away.
  switchRole(role: { role_id: number; role_name: string }): void {
    this.showRoleMenu = false;

    if (role.role_name === this.currentRoleName) {
      return;
    }

    this.service.switchRole(role.role_id).subscribe({
      next: (res) => {
        sessionStorage.setItem('authToken', res.access_token);
        sessionStorage.setItem('roleId', String(role.role_id));
        sessionStorage.setItem('roleName', role.role_name);
        this.currentRoleName = role.role_name;

        // Role-gated UI/data needs a full reload to reflect the new role
        // cleanly rather than trying to patch every open view in place —
        // but a bare reload alone isn't enough: it would just re-hit
        // whatever route the user happened to be on under the OLD role
        // (e.g. switching OUT of Admin while sitting on /admin would
        // reload straight back into a now-unauthorized admin route, and
        // switching INTO Admin from /dashboard would just reload
        // /dashboard instead of surfacing the admin page). Land on the
        // page appropriate for the role just switched to, then reload.
        const targetPath = role.role_name.toLowerCase() === 'admin' ? 'admin' : 'dashboard';
        const baseUrl = window.location.href.split('#')[0];
        window.location.href = `${baseUrl}#/${targetPath}`;
        window.location.reload();
      },
      error: (err) => {
        console.error('Role switch failed:', err);
      },
    });
  }

  // Click-outside handling: close the role dropdown if the user clicks
  // anywhere other than the switcher button/menu itself. Since the
  // dropdown only exists in the DOM (*ngIf) while availableRoles.length
  // > 1, roleSwitcherRef itself is only populated in that same case —
  // if it's undefined there's nothing to be "inside" of anyway.
  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent): void {
    if (!this.showRoleMenu) {
      return;
    }
    const switcherEl = this.roleSwitcherRef?.nativeElement;
    if (!switcherEl || !switcherEl.contains(event.target as Node)) {
      this.showRoleMenu = false;
    }
  }

  ngOnInit(): void {
    const username = sessionStorage.getItem('userName') || '';
    this.userInitials = this.getInitials(username);
    // Cached-only on init — no network call until the dropdown is opened.
    this.loadRoles();
    this.VERSION = VERSION.version;
    console.log('App Version:', this.VERSION);
  }

  getInitials(name: string): string {
    return name
      .trim()
      .split(' ')
      .filter(word => word.length > 0)
      .map(word => word[0].toUpperCase())
      .slice(0, 2)
      .join('');
  }

  onSearch(): void {
    console.log('Search:', this.searchQuery);
  }

  onNotificationClick(): void {
    console.log('Notifications clicked');
  }
  showLogoutPopup = false;
  onLogout(): void {
  this.showLogoutPopup = true;
  }

  confirmLogout(): void {
    this.showLogoutPopup = false;
    sessionStorage.clear();
    console.log('Logout clicked');
    this.router.navigate(['/logout']);
  }

  cancelLogout(): void {
    this.showLogoutPopup = false;
  }

  // onLogout(): void {
  //   sessionStorage.clear();
  //   console.log('Logout clicked');
  //   this.router.navigate(['/logout']);
  // }
  
}