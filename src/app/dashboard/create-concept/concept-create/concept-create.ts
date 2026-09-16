import { Component, OnInit, ViewChild, ElementRef, ChangeDetectorRef, OnDestroy } from '@angular/core';
import { CommonModule, DatePipe } from '@angular/common';
import { ReactiveFormsModule, FormBuilder, FormGroup, Validators } from '@angular/forms';
import { FormsModule } from '@angular/forms';
import { LayoutComponent } from "../../../layout/layout/layout";
import { ToastrService } from 'ngx-toastr';
import { Service } from '../../../dashboard/service';
import { DomSanitizer, SafeHtml, SafeResourceUrl } from '@angular/platform-browser';
import { ActivatedRoute, Router } from '@angular/router';
import * as mammoth from 'mammoth';
import * as XLSX from 'xlsx';
import { Subscription } from 'rxjs';


interface User { id: string; name: string; }
interface Owner { id: string; name: string; initials: string; avatarColor: string; department: string; }
interface Tab { key: string; label: string; }

interface OwnerCandidate {
  id: string;
  name: string;
  department: string;
  initials: string;
  avatarColor: string;
}

interface DevNote {
  id: string;
  author: string;
  initials: string;
  avatarBg: string;
  time: string;
  text: string;
  RoleName: string;
  /** true once this note has actually been saved to the backend (loaded
   *  from the API, or confirmed sent in a previous submit). Only notes
   *  with persisted === false get included in the next submit's
   *  developmentNotes payload — otherwise every Update re-sends (and the
   *  backend re-inserts) every note that already exists, duplicating them
   *  on every save. */
  persisted: boolean;
}

interface AttachFile {
  id: string;             // stable identity for progress updates — object refs change on each tick
  name: string;
  size: number;
  progress: number;
  file: File;
  downloadUrl?: string;   // set when restored from a saved concept (for fetching real bytes on view)
  /** Backend AttachmentId — set only for files restored from a saved
   *  concept. Undefined for a file just picked this session, since it
   *  has no backend record yet. Drives whether removeAttachment() needs
   *  to call the delete API or can just drop it from the local array. */
  attachmentId?: number;
}

interface SupportingDoc {
  name: string;
  sourceurl: string;
  pdfLocation: string;
  uploadProgress: number;
  file: File | null;
  downloadUrl?: string;   // set when restored from a saved concept — GET /api/download-attachment/{id}
  restored?: boolean;     // true when this card was filled from a saved concept, not just uploaded this session
  /** Controls the "Uploaded successfully" banner. Left undefined for
   *  restored docs so their banner (which reads "Uploaded" instead) stays
   *  visible permanently. Explicitly set true by simulateDocUpload() once
   *  a freshly-picked file hits 100%, then flipped to false 5s later so
   *  the banner auto-dismisses without touching the rest of the card. */
  successBannerVisible?: boolean;
  /** Last known FileName/FileSize from the backend for this slot, captured
   *  on restore. `file` for a restored doc is just a zero-byte placeholder
   *  (see patchSupportingDocs), so when the slot is resubmitted without a
   *  newly-picked file, these are what tell the backend "this slot's
   *  existing file is still intact" — otherwise the update looks
   *  file-less and the backend can wipe the file association on its end
   *  (see onDocSubmit's resolveDocFileMeta()). */
  originalFileName?: string;
  originalFileSize?: number;
  /** Backend AttachmentId for this slot — set only when restored from a
   *  saved concept. Undefined for a slot added/filled this session that
   *  was never submitted, since there's no backend record to delete yet.
   *  Drives whether removeSupportingDoc() needs to call the delete API. */
  attachmentId?: number;
  /** STABLE slot identity, mirrors the backend's ConceptAttachments.DocIndex
   *  column. Assigned exactly once per slot — either read back from the
   *  backend on restore, or handed out by allocateDocIndex() the moment a
   *  brand-new slot is created — and NEVER recomputed from the slot's
   *  position in the `supportingDocs` array. Array position shifts every
   *  time a doc is deleted (Array.filter), but this must not, otherwise a
   *  later doc silently inherits an earlier (now-deleted) doc's DocIndex
   *  and the backend's UPDATE ... WHERE DocIndex = ? can deactivate the
   *  wrong attachment row. See onDocSubmit() / resolveDocFileMeta(). */
  docIndex: number;
}

interface SheetData {
  sheetName: string;
  headers: string[];
  rows: any[][];
}

interface LatestConceptItem {
  /** Stable anchor id — PK of ConceptKeys/Concepts, never changes across
   *  edits. This is what must be used for routing and for this.conceptId,
   *  never CurrentConceptId (see onSelectLatestConcept). Requires the
   *  backend list endpoint (getConceptsByUserId) to select this column. */
  ConceptId: string;
  /** Display/version id (…_D001 → …_D002 → …_D003, or the finalized
   *  production id). Changes on every dev edit — label/display only,
   *  never send this back to the backend as concept_id. */
  CurrentConceptId: string;
  ConceptName: string;
  DevelopmentStatus: string;
  CreatedDate: string;
  statusClass?: string;
  RecordType?: string;
  isDraft?: boolean;
  /** Assignment fields used by the "Assigned to me" filter. The
   *  /api/latest-updates response only ever returns the display names
   *  (IdeationRequestor / DataScienceProgrammer) — no id columns — so
   *  isAssignedToCurrentUser() matches purely on username. */
  IdeationRequestor?: string;
  DataScienceProgrammer?: string;
  IdeationRequestorId?: Number;
  DataScienceProgrammerId?: Number;
}

/** One row from GET /api/concepts/{concept_id}/history, already
 *  pre-shaped server-side for direct UI binding — see concept_history.py.
 *  dot_color/title/timestamp drive the collapsed timeline row;
 *  old_value/new_value/changed_by/role/reason drive the click-through
 *  detail popup, using the exact same object (no second API call). */
interface ActivityHistoryItem {
  activity_id: number;
  activity_type: string;
  dot_color: 'red' | 'green' | 'blue' | 'yellow' | 'gray' | 'purple';
  title: string;
  timestamp: string;
  old_value: string;
  new_value: string;
  changed_by: string;
  role: string;
  reason: string;
}

interface IdValueOption { id: number; value: string; }
interface ClientOption { client_id: string; client_name: string; }
interface MasterConceptOption { master_id: string; concept_name: string; }
interface ReviewTypeOption { review_type: string; description: string; }
interface ClaimTypeOption { claim_type: string; description: string; }
interface IdeationRequestorOption { id: number; name: string; role_name: string; }
interface DataScienceProgrammerOption { id: number; name: string; role_name?: string; }
interface clientApprovalstatusOption { id: number; value: string }

type AttachCategory = 'specs' | 'table' | 'other' | 'approval';

@Component({
  selector: 'app-concept-create',
  standalone: true,
  imports: [CommonModule, ReactiveFormsModule, FormsModule, DatePipe, LayoutComponent],
  templateUrl: './concept-create.html',
  styleUrls: ['./concept-create.css']
})
export class ConceptCreateComponent implements OnInit, OnDestroy {
  @ViewChild('specsInput') specsInputRef!: ElementRef<HTMLInputElement>;
  @ViewChild('tableInput') tableInputRef!: ElementRef<HTMLInputElement>;
  @ViewChild('otherInput') otherInputRef!: ElementRef<HTMLInputElement>;
  @ViewChild('approvalInput') approvalInputRef!: ElementRef<HTMLInputElement>;
  @ViewChild('fileInput') fileInput!: ElementRef<HTMLInputElement>;
  @ViewChild('docFileInput') docFileInputRef!: ElementRef<HTMLInputElement>;
  @ViewChild('notesListRef') notesListRef!: ElementRef<HTMLDivElement>;
  @ViewChild('tabContentRef') tabContentRef!: ElementRef<HTMLDivElement>;
  @ViewChild('activityLegendWrap') activityLegendWrapRef?: ElementRef<HTMLElement>;
  private pendingDocIndex: number | null = null;

  // ── Mode ───────────────────────────────────────────────────────────────
  isEditMode = false;
  pageLoading = false;

  // ── Activity History (right panel card + detail popup) ──────────────────
  activityHistory: ActivityHistoryItem[] = [];
  activityHistoryLoading = false;
  /** Collapsed card shows only the most recent N entries; "View full
   *  history" flips this to render the complete list in place. */
  showFullActivityHistory = false;
  readonly activityHistoryPreviewCount = 4;
  /** Non-null while the detail popup (Image 2) is open — bound directly
   *  from whichever ActivityHistoryItem the user clicked, so no second
   *  request is needed to populate the popup. */
  selectedActivity: ActivityHistoryItem | null = null;
  /** Controls the small "what do the dot colors mean?" legend popover
   *  next to the card header (the ⓘ icon). */
  showActivityLegend = false;
  /** Viewport (fixed-position) coordinates for the legend popover, computed
   *  right before it opens (see positionActivityLegend()). Using `fixed`
   *  positioning instead of `absolute` means the popover is placed relative
   *  to the viewport rather than to the scrollable .right-panel, so it can
   *  no longer be clipped by that panel's overflow, nor visually buried
   *  under the center panel. */
  activityLegendPos: { top: number; left: number } = { top: 0, left: 0 };
  /** True while a copy-concept API call is in flight — disables the Copy
   *  button and swaps its label so a slow response can't be double-fired. */
  copyingConcept = false;
  /** Built entirely from the history response's dot_color_map (see
   *  loadActivityHistory -> buildActivityDotLegend) — no hardcoded rows
   *  here, so the legend only ever shows colors/labels the backend has
   *  actually confirmed, and picks up any future color the backend adds
   *  without a code change. Empty until that response arrives (e.g. no
   *  concept loaded yet); the popover shows a "not available yet" line in
   *  that state instead of a hardcoded row set. */
  activityDotLegend: { color: string; colorLabel: string; label: string }[] = [];

  // The last-PERSISTED Concept Name (set from patchForm() on load/reload),
  // deliberately kept separate from the live form.conceptName value. The
  // Client Approval tab's clientConceptName must only ever mirror what's
  // actually been saved — an in-progress, not-yet-saved edit to Concept
  // Name on the Development tab must NOT show up over on the Client
  // Approval tab until that edit is actually saved/submitted. See
  // ensureClientConceptName().
  private savedConceptName = '';

  // ── Meta ──────────────────────────────────────────────────────────────
  // conceptId is the STABLE ANCHOR — PK of ConceptKeys/Concepts on the
  // backend, never changes across edits. This is what goes in the route
  // and in every submit's concept_id field.
  conceptId = '';
  // displayConceptId is the version/display id (…_D001 → …_D002 → …_D003,
  // or the finalized production id) — for showing to the user only.
  // NEVER send this back to the backend as concept_id.
  displayConceptId = '';
  createdDate = new Date();
  updatedDate = new Date();

  // ── Section completion flags ─────────────────────────────────────────
  // Each one flips to 1 only when its own tab's Submit button is clicked,
  // and is sent to the backend inside `metadata` in submitConcept().
  developmentCompleted: 0 | 1 = 0;
  clientApprovalCompleted: 0 | 1 = 0;
  supportingDocumentsCompleted: 0 | 1 = 0;
  isDraftConcept = false;

  /** Concept Status card's Client Approval sub-label. `clientApprovalCompleted`
   *  only tells you the step has been submitted at all — it doesn't say
   *  whether that submission is sitting with the client, approved, or
   *  rejected. The actual workflow state lives in the latest
   *  client_approvals row's ClientApprovalStatus (e.g. "Submitted",
   *  "Approved", "Rejected") and patchClientApproval() already copies
   *  that into the clientApprovalStatus form control, so read it from
   *  there rather than re-deriving a generic yes/no string. Falls back
   *  to a generic label only if the backend hasn't sent a status yet. */
  get clientApprovalStatusLabel(): string {
    const status = (this.form.get('clientApprovalStatus')?.value ?? '').toString().trim();
    if (status) return status;
    return this.clientApprovalCompleted ? 'Submitted' : 'Not Submitted';
  }

  /** Shown when the user clicks Client Approval / Supporting Document
   *  while the concept is still a draft. Tells them why the click did
   *  nothing instead of relying on the title tooltip alone. */
  showDraftLockBanner = false;
  private draftLockBannerTimer: ReturnType<typeof setTimeout> | null = null;

  /** Shown briefly whenever a Production-locked concept loads — same
   *  flash-in / auto-dismiss pattern as showDraftLockBanner above rather
   *  than sitting on screen indefinitely. See flashProductionLockBanner(). */
  showProductionLockBanner = false;
  private productionLockBannerTimer: ReturnType<typeof setTimeout> | null = null;


  // ── Navigation ────────────────────────────────────────────────────────
  searchQuery = '';
  leftPanelCollapsed = false;

  /** "Assigned to me" toggle for the Latest Updates list — when true,
   *  filteredConcepts is narrowed to concepts where the logged-in user
   *  (sessionStorage userId / userName) is the IdeationRequestor or the
   *  DataScienceProgrammer. Combines with the text search below. */
  assignedToMeOnly = false;

  /** Matches a Latest-Updates row against the current session user for
   *  the IdeationRequestor / DataScienceProgrammer roles. The backend
   *  (/api/latest-updates) only ever returns display names for these —
   *  never ids — so this is a case-insensitive name match against
   *  sessionStorage's userName. */
  private isAssignedToCurrentUser(item: LatestConceptItem): boolean {
    const sessionUserName = (sessionStorage.getItem('userName') || '').trim().toLowerCase();
    const sessionUserid = Number(sessionStorage.getItem('userId') || '');
    if (!sessionUserName || !sessionUserid) {
      return false;
    }

    const requestorName = (item.IdeationRequestor || '').trim().toLowerCase();
    const requestorid = (item.IdeationRequestorId || '');
    const programmerName = (item.DataScienceProgrammer || '').trim().toLowerCase();
    const programmerid = (item.DataScienceProgrammerId || '');

    return requestorName === sessionUserName || programmerName === sessionUserName || requestorid === sessionUserid || programmerid === sessionUserid;

  }

  get filteredConcepts(): LatestConceptItem[] {
    let list = this.latestConcepts;

    if (this.assignedToMeOnly) {
      list = list.filter(c => this.isAssignedToCurrentUser(c));
    }

    const q = this.searchQuery.trim().toLowerCase();
    if (!q) return list;
    return list.filter(c =>
      c.ConceptName?.toLowerCase().includes(q) ||
      c.CurrentConceptId?.toLowerCase().includes(q) ||
      c.DevelopmentStatus?.toLowerCase().includes(q)
    );
  }
  activeTab = 'development';

  tabs: Tab[] = [
    { key: 'development', label: 'Concept Development' },
    { key: 'approval', label: 'Client Approval' },
    { key: 'documents', label: 'Supporting Document' }
  ];

  // ── Form ──────────────────────────────────────────────────────────────
  form!: FormGroup;

  // ── Reference data ────────────────────────────────────────────────────
  users: User[] = [
    { id: 'u1', name: 'Alice Johnson' },
    { id: 'u2', name: 'Bob Martinez' },
    { id: 'u3', name: 'Carol Singh' },
    { id: 'u4', name: 'David Chen' }
  ];

  owners: Owner[] = [];
  uploadedFiles: File[] = [];

  // ── Ownership dropdown ────────────────────────────────────────────────
  showOwnerDropdown = false;
  ownerSearch = '';

  ownerCandidates: OwnerCandidate[] = [
    { id: 'p1', name: 'Alice Johnson', initials: 'AJ', avatarColor: '#6366f1', department: 'Data Science' },
    { id: 'p2', name: 'Bob Martinez', initials: 'BM', avatarColor: '#8b5cf6', department: 'Engineering' },
    { id: 'p3', name: 'Carol Singh', initials: 'CS', avatarColor: '#ec4899', department: 'Product' },
    { id: 'p4', name: 'David Chen', initials: 'DC', avatarColor: '#f59e0b', department: 'Analytics' },
    { id: 'p5', name: 'Emma Williams', initials: 'EW', avatarColor: '#10b981', department: 'QA' },
    { id: 'p6', name: 'Frank Torres', initials: 'FT', avatarColor: '#3b82f6', department: 'Data Science' },
    { id: 'p7', name: 'Grace Kim', initials: 'GK', avatarColor: '#ef4444', department: 'Engineering' }
  ];

  latestConcepts: LatestConceptItem[] = [];
  latestUpdatesLoading = false;

  // ── Master data (dropdown sources) ────────────────────────────────────
  developmentStatusOptions: IdValueOption[] = [];
  priorityOptions: IdValueOption[] = [];
  clientOptions: ClientOption[] = [];
  masterConceptOptions: MasterConceptOption[] = [];
  reviewTypeOptions: ReviewTypeOption[] = [];
  claimTypeOptions: ClaimTypeOption[] = [];
  ideationRequestorOptions: IdeationRequestorOption[] = [];
  dataScienceProgrammerOptions: DataScienceProgrammerOption[] = [];
  clientApprovalstatusOptions: clientApprovalstatusOption[] = [];
  masterDataLoading = false;

  // ── Date constraints ──────────────────────────────────────────────────
  /** Today's date in yyyy-MM-dd, used as the [min] bound on QA Schedule /
   *  Production Schedule / Submitted To Client On date pickers so past
   *  dates aren't selectable. */
  minSelectableDate = new Date().toISOString().split('T')[0];

  // ── Locale-independent date display (MM/DD/YYYY) ────────────────────
  /**
   * Native <input type="date"> renders its typed segments using the
   * OS/browser locale (MM/DD/YYYY on one machine, DD/MM/YYYY on another) —
   * there's no HTML attribute that overrides this. QA Schedule, Production
   * Schedule, and Submitted To Client On are instead rendered as plain
   * text inputs that WE format, always as MM/DD/YYYY, paired with a
   * hidden native date input (per field) used only to pop the native
   * calendar via showPicker(). The FormControl itself still always holds
   * a plain 'yyyy-MM-dd' string, exactly as before, so validDateRange /
   * notPastDate and every submit payload are unaffected.
   */

  /** Holds whatever the user actually typed for a date field once it's
   *  been left in a state that doesn't resolve to a real calendar date
   *  (e.g. "13/45/2026"). The FormControl itself is never set to this —
   *  it must keep holding a plain 'yyyy-MM-dd' string or '' (see note
   *  above) — so this is what getDateDisplay() falls back to instead of
   *  silently re-deriving blank/stale text from the control on the next
   *  change-detection pass. Cleared as soon as the field resolves to a
   *  valid date or is genuinely emptied out. */
  private invalidDateText: { [controlName: string]: string } = {};

  /** Bound to the visible text box's [value]; always MM/DD/YYYY. */
  getDateDisplay(controlName: string): string {
    if (this.invalidDateText[controlName] !== undefined) {
      return this.invalidDateText[controlName];
    }
    return this.isoToDisplayDate(this.form.get(controlName)?.value);
  }

  /** User typed into the visible text box. */
  onDateTextInput(event: Event, controlName: string): void {
    if (this.form.get(controlName)?.disabled) return;
    const input = event.target as HTMLInputElement;
    const formatted = this.autoFormatDateInput(input.value);
    input.value = formatted;
    const control = this.form.get(controlName);
    if (formatted.length === 10) {
      const iso = this.displayDateToIso(formatted);
      if (iso) {
        delete this.invalidDateText[controlName];
        control?.markAsDirty();
        control?.setValue(iso);
      } else {
        // A full 10 characters that still isn't a real calendar date
        // (e.g. pasted "13/45/2026", or "46/54/6546"). Flag it now
        // rather than waiting for blur, so the field shows as invalid
        // immediately.
        this.invalidDateText[controlName] = formatted;
        control?.markAsDirty();
        // Spreading the control's PRIOR errors here is what caused the
        // "Submitted To Client On is required." message to show up
        // alongside "Please enter a valid date" — required had already
        // been computed true back when the box was still empty, and
        // spreading {...control?.errors} onto the new error object
        // carried that stale flag forward even though the user had
        // since typed 10 characters into it. setErrors() bypasses the
        // validators array entirely, so nothing ever told Angular to
        // drop `required` once the field stopped being empty. Explicitly
        // dropping it here means only the one relevant message —
        // invalidDateRange — shows for an unparseable-but-non-empty date.
        const { required, ...otherErrors } = control?.errors || {};
        control?.setErrors({ ...otherErrors, invalidDateRange: true });
      }
    } else if (formatted.length > 0) {
      // Still mid-typing — keep the box showing exactly what's been
      // typed so far instead of letting getDateDisplay() re-derive it
      // from the (unchanged) control value on the next change-detection
      // pass, which would wipe out an in-progress entry.
      this.invalidDateText[controlName] = formatted;
    } else {
      // Box is genuinely empty (select-all + delete, backspaced to
      // nothing, etc) DURING typing — not just discovered later on blur.
      // Previously this branch only cleared the invalidDateText override
      // and left the FormControl untouched, which caused two visible
      // bugs at once:
      //   1. getDateDisplay() falls back to isoToDisplayDate(control.value)
      //      whenever there's no invalidDateText override — so on the very
      //      next change-detection tick the box "snapped back" to the old
      //      date instead of staying empty, because the control's actual
      //      value was never cleared.
      //   2. If a prior keystroke had left stale errors on the control via
      //      the manual control?.setErrors({ invalidDateRange: true }) call
      //      above (or notPastDate from an earlier value), those errors
      //      were never recomputed — setErrors() bypasses the validators
      //      array entirely, so nothing here ever told Angular to re-run
      //      validDateRange/notPastDate/required and discover they now all
      //      pass for an empty value. The old invalidDateRange/pastDate
      //      flags just sat there, so Submit could show TWO error toasts
      //      (e.g. "Please enter a valid QA Schedule date." AND the
      //      Production Schedule one) for fields that looked empty on
      //      screen but were still internally "invalid".
      // Actually clearing the control here — same as the equivalent branch
      // in onDateTextBlur — fixes both: setValue('') re-runs the real
      // validators, which return null for an empty value (see
      // validDateRange/notPastDate), wiping any stale manually-set errors,
      // and the box now has nothing stale to snap back to.
      delete this.invalidDateText[controlName];
      control?.markAsDirty();
      control?.setValue('');
    }
  }

  /** Leaving the text box: reconcile — a valid date commits, a fully
   *  emptied box clears the control (not just the text), and anything
   *  left behind that isn't a real calendar date (e.g. "13/45/2026" or
   *  a truncated entry) is now kept on screen and flagged invalid
   *  instead of being silently discarded.
   *
   *  Previously the "incomplete/invalid" branch reverted the visible
   *  text back to whatever the control still held (often '') and did
   *  nothing to the control's validity. Clicking Update immediately
   *  after typing a bad date blurs the field first, so that revert ran
   *  right before the click handler read the form — the box went blank
   *  in the same tick, form.valid was still true (the control's actual
   *  value never changed), and the update proceeded with the success
   *  popup even though the typed date was never saved. */
  onDateTextBlur(event: Event, controlName: string): void {
    if (this.form.get(controlName)?.disabled) return;
    const input = event.target as HTMLInputElement;
    const control = this.form.get(controlName);
    const iso = this.displayDateToIso(input.value);
    if (input.value && iso) {
      delete this.invalidDateText[controlName];
      control?.markAsDirty();
      control?.setValue(iso);
    } else if (input.value) {
      // Doesn't resolve to a real calendar date. Keep it visible (don't
      // wipe what the user typed) and flag the control invalid so the
      // Update/Submit handlers' existing hasError('invalidDateRange')
      // checks catch it and block the save, instead of quietly letting
      // an unrelated old value (or nothing) through.
      //
      // Deliberately NO toast here — the visual ng-invalid state on the
      // input (see template) is the immediate feedback. A toast on blur
      // AND another toast on Submit for the exact same invalidDateRange
      // state was firing twice back-to-back: clicking Submit blurs the
      // focused field first (native browser behavior), which ran this
      // handler and toasted once, then onSubmit's own hasError() check
      // ran right after and toasted again for the same problem. Submit
      // is now the single place this error surfaces as a toast.
      this.invalidDateText[controlName] = input.value;
      control?.markAsDirty();
      // Same reasoning as the onDateTextInput equivalent above: drop any
      // stale `required` flag before merging in invalidDateRange, so a
      // non-empty-but-unparseable date (e.g. "46/54/6546") shows only
      // "Please enter a valid date" instead of that message AND
      // "...is required." at the same time.
      const { required, ...otherErrors } = control?.errors || {};
      control?.setErrors({ ...otherErrors, invalidDateRange: true });
    } else {
      // Box is genuinely empty (user selected all + deleted, backspaced
      // to nothing, etc). Previously this fell through both branches
      // above and left the FormControl holding its last committed ISO
      // date — so a required-date field like Submitted To Client On
      // still read as "filled in" to isRequiredFieldMissing()/Validators
      // .required even though the visible box was blank, letting
      // Client Approval (and any other required date) submit
      // successfully with no date shown. Clear the control to match
      // what's actually on screen.
      delete this.invalidDateText[controlName];
      control?.markAsDirty();
      control?.setValue('');
    }
    control?.markAsTouched();
  }

  /** User picked a date from the native calendar popup. */
  onDateNativeChange(event: Event, controlName: string): void {
    if (this.form.get(controlName)?.disabled) return;
    const input = event.target as HTMLInputElement;
    const control = this.form.get(controlName);
    // Clear any leftover invalid-text override from a prior bad manual
    // entry (e.g. "12/34/5656") — otherwise getDateDisplay() keeps
    // returning that stale string instead of the newly picked date,
    // even though the control itself now holds the correct value. This
    // is why the box looked stuck on the old invalid date while the
    // correct one was silently going through on submit.
    delete this.invalidDateText[controlName];
    control?.markAsDirty();
    control?.setValue(input.value || '');
    control?.markAsTouched();
  }

  /** Opens the hidden native <input type="date"> calendar for a given
   *  field. Pass the template reference variable of the hidden input. */
  openDatePicker(nativeInput: HTMLInputElement): void {
    if (!nativeInput || nativeInput.disabled) return;
    const anyNative = nativeInput as any;
    if (typeof anyNative.showPicker === 'function') {
      try {
        anyNative.showPicker();
        return;
      } catch {
        // fall through to focus/click fallback below
      }
    }
    nativeInput.focus();
    nativeInput.click();
  }

  /** As digits are typed, auto-inserts '/' after MM and DD. */
  private autoFormatDateInput(raw: string): string {
    const digits = raw.replace(/\D/g, '').slice(0, 8); // MMDDYYYY
    let out = '';
    if (digits.length > 0) out += digits.slice(0, 2);
    if (digits.length >= 3) out += '/' + digits.slice(2, 4);
    else if (digits.length > 2) out += '/' + digits.slice(2);
    if (digits.length >= 5) out += '/' + digits.slice(4, 8);
    return out;
  }

  /** 'MM/DD/YYYY' -> 'yyyy-MM-dd', or '' if not a real, complete date. */
  private displayDateToIso(display: string): string {
    const match = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(display);
    if (!match) return '';
    const month = parseInt(match[1], 10);
    const day = parseInt(match[2], 10);
    const year = parseInt(match[3], 10);
    if (month < 1 || month > 12 || day < 1 || day > 31) return '';
    const date = new Date(year, month - 1, day);
    // Guard against overflow dates like 02/31/2026 rolling into March.
    if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) {
      return '';
    }
    const mm = String(month).padStart(2, '0');
    const dd = String(day).padStart(2, '0');
    return `${year}-${mm}-${dd}`;
  }

  /** 'yyyy-MM-dd' -> 'MM/DD/YYYY', or '' if empty/unparseable. */
  private isoToDisplayDate(iso: string | null | undefined): string {
    const match = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || '');
    if (!match) return '';
    const [, year, month, day] = match;
    return `${month}/${day}/${year}`;
  }



  // ── Role-Based Access Control ─────────────────────────────────────────
  /** Numeric role ids, assigned by the backend/auth layer and stored in
   *  sessionStorage as 'roleId':
   *    1 = Ideation Requestor
   *    2 = Data Science Programmer
   *    3 = Manager
   *    4 = QA
   *    5 = Viewer
   *    6 = Operations
   *  Every permission check below is driven off this numeric id, NOT the
   *  display name — role names can be renamed/localized on the backend
   *  without silently breaking access control here. */
  private static readonly ROLE_IDEATION_REQUESTOR = 1;
  private static readonly ROLE_DATA_SCIENCE_PROGRAMMER = 2;
  private static readonly ROLE_MANAGER = 3;
  private static readonly ROLE_QA = 4;
  private static readonly ROLE_VIEWER = 5;
  private static readonly ROLE_OPERATIONS = 6;

  currentUserRoleId: number = Number(sessionStorage.getItem('roleId')) || 0;

  /** Display-only — kept purely in case any template wants to show the
   *  role's name somewhere. NEVER used in a permission check anymore;
   *  every get canX()/isReadOnly below reads currentUserRoleId instead. */
  currentUserRole: string = sessionStorage.getItem('roleName') ?? '';
  // console.log('Current user role id:', this.currentUserRoleId);

  /** True for a Manager login. Used specifically by the Production-lock
   *  logic below — Manager is the one role allowed to still touch the
   *  Development Status dropdown once a concept has reached Production. */
  get isManager(): boolean {
    return this.currentUserRoleId === ConceptCreateComponent.ROLE_MANAGER;
  }

  /** The concept's own PERSISTED Development Status, as last returned by
   *  the backend — set in patchForm() and reset in buildForm(). This is
   *  deliberately NOT the live (possibly unsaved) value sitting in the
   *  developmentStatus form control, so that a Manager picking a new
   *  status from the dropdown doesn't instantly lock/unlock the rest of
   *  the form out from under them before that change is actually saved. */
  private currentConceptStatus: string = '';

  /** Once a concept's Development Status is "Production" the record is
   *  treated as finalized: every field on every tab is force-locked
   *  read-only for every role. The one exception is Manager, who may
   *  still change the Development Status dropdown itself (e.g. to move
   *  it forward/back or correct a mistake) — nothing else. Read by
   *  applyRoleRestrictions(), applyAllowedStatusFilter(), the
   *  attachment/notes/approval gates below, and the template. */
  get isProductionLocked(): boolean {
    return this.currentConceptStatus === 'Production';
  }

  get isSupersededLocked(): boolean {
    return this.currentConceptStatus === 'Superseded';
  }

  get isRecordLocked(): boolean {
    return this.isProductionLocked || this.isSupersededLocked;
  }
  /** Full edit + submit rights. Ideation Requestor, QA, Manager. */
  get canFullEdit(): boolean {
    return [
      ConceptCreateComponent.ROLE_IDEATION_REQUESTOR,
      ConceptCreateComponent.ROLE_QA,
      ConceptCreateComponent.ROLE_MANAGER,
    ].includes(this.currentUserRoleId);
  }
  /** Data Science Programmer — limited field edit + submit, no approval tab. */
  get canDSEdit(): boolean {
    return this.currentUserRoleId === ConceptCreateComponent.ROLE_DATA_SCIENCE_PROGRAMMER;
  }

  /** Any role that can write something (not pure read-only). */
  get canEdit(): boolean {
    return this.canFullEdit || this.canDSEdit;
  }

  /** Gates the Development Notes input specifically. Identical to
   *  canEdit normally. While Production-locked, notes stay off-limits
   *  for everyone EXCEPT Manager — a Manager changing the Development
   *  Status is exactly the case where they'd want to leave a note
   *  explaining the change, so Manager keeps this one even though every
   *  other field (attachments, approvals, supporting docs, etc.) stays
   *  locked. */
  get canEditNotes(): boolean {
    if (this.isRecordLocked) return this.isProductionLocked && this.isManager;
    return this.canEdit;
  }

  /** Gates the Save-as-Draft / Update buttons specifically. Identical to
   *  canEdit normally, but while the concept is Production-locked every
   *  other field is disabled (see applyRoleRestrictions()), so only a
   *  Manager — saving nothing but their own Development Status change
   *  (and, per canEditNotes above, an optional note about it) — is
   *  allowed to still click through. Everyone else loses the button
   *  entirely once locked, even if their role could otherwise edit. */
  get canSubmitForm(): boolean {
    if (this.isRecordLocked) return this.isProductionLocked && this.isManager;
    return this.canEdit;
  }

  /** Gates Save-as-Draft and the "Create Concept" submit action
   *  specifically (as opposed to "Update"). Data Science Programmer
   *  works existing concepts assigned to them and never originates new
   *  ones — canCreateConcept already hides the "+ Add New Concept" entry
   *  point for that role, but this closes the gap for a DS Programmer
   *  who lands on a blank (not-yet-created) concept form some other way:
   *  both buttons stay disabled for them until the concept already
   *  exists (developmentCompleted), at which point "Update" is governed
   *  by canSubmitForm as normal. */
  get canSaveOrCreate(): boolean {
    if (this.canDSEdit && !this.developmentCompleted) return false;
    return this.canSubmitForm;
  }

  /** Only full-edit roles (Ideation Requestor, QA, Manager) may create a
   *  brand-new concept. Data Science Programmer works existing concepts
   *  assigned to them but doesn't originate new ones, and read-only
   *  roles obviously can't either — gates the "+ Add New Concept" button. */
  get canCreateConcept(): boolean {
    return this.canFullEdit;
  }

  /** Only Manager and full-edit roles may submit Client Approval —
   *  further gated to nothing at all once the concept is Production-
   *  locked, Manager included (Manager's only remaining edit right at
   *  that point is the Development Status dropdown itself). */
  get canSubmitApproval(): boolean {
    if (this.isRecordLocked) return false;
    return [
      ConceptCreateComponent.ROLE_IDEATION_REQUESTOR,
      ConceptCreateComponent.ROLE_QA,
      ConceptCreateComponent.ROLE_MANAGER,
    ].includes(this.currentUserRoleId);
  }

  /** For Manager logins ONLY, both the Ideation Requestor and Data
   *  Science Programmer dropdowns show the full combined pool of people
   *  from BOTH master-data lists (ideation_requestors +
   *  datascience_programmers) — same source as GET /api/master-data —
   *  each labeled "<name> - <role_name>", so a Manager can assign either
   *  role to anyone regardless of which list they originally came from.
   *  Every other role keeps seeing only its own single-role list. */
  // get ideationRequestorDropdownOptions(): IdeationRequestorOption[] {
  //   return this.currentUserRoleId === ConceptCreateComponent.ROLE_MANAGER
  //     ? this.mergeUserOptionLists(this.ideationRequestorOptions, this.dataScienceProgrammerOptions)
  //     : this.ideationRequestorOptions;
  // }
  get ideationRequestorDropdownOptions(): IdeationRequestorOption[] {
    // Always return only the ideation_requestors list
    return this.ideationRequestorOptions;
  }

  // get dataScienceProgrammerDropdownOptions(): DataScienceProgrammerOption[] {
  //   return this.currentUserRoleId === ConceptCreateComponent.ROLE_MANAGER
  //     ? this.mergeUserOptionLists(this.dataScienceProgrammerOptions, this.ideationRequestorOptions)
  //     : this.dataScienceProgrammerOptions;
  // }
  get dataScienceProgrammerDropdownOptions(): DataScienceProgrammerOption[] {
    // Always return only the datascience_programmers list
    return this.dataScienceProgrammerOptions;
  }

  /** Combines two option lists into one, de-duplicated by id — primary
   *  list's entries win when the same id appears in both (e.g. someone
   *  like Praveen Choudhary in the sample data, who holds both roles). */
  private mergeUserOptionLists(
    primary: { id: number; name: string; role_name?: string }[],
    secondary: { id: number; name: string; role_name?: string }[]
  ): any[] {
    const merged = [...primary];
    const seenIds = new Set(primary.map(u => Number(u.id)));
    for (const u of secondary) {
      if (!seenIds.has(Number(u.id))) {
        merged.push(u);
        seenIds.add(Number(u.id));
      }
    }
    return merged;
  }

  /** Pure read-only roles. */
  get isReadOnly(): boolean {
    return [
      ConceptCreateComponent.ROLE_OPERATIONS,
      ConceptCreateComponent.ROLE_VIEWER,
    ].includes(this.currentUserRoleId);
  }

  /** Gates every file upload/delete action (Attachments AND Supporting
   *  Documents), not just the form fields. Previously only the form
   *  controls were disabled for read-only/unauthorized roles — the
   *  attach/upload/remove buttons and their underlying handlers had no
   *  permission check at all, so a read-only user could still add or
   *  delete files even though every other field on the page was locked.
   *  'approval' attachments are gated by canSubmitApproval (Client
   *  Approval has its own, narrower permission set) — every other
   *  category (specs/table/other) and Supporting Documents are gated by
   *  the general canEdit. Public so the template can also use it to
   *  hide/disable the relevant buttons, not just block the click. */
  canManageAttachments(cat: AttachCategory): boolean {
    if (this.isReadOnly || this.isRecordLocked) return false;
    return cat === 'approval' ? this.canSubmitApproval : this.canEdit;
  }

  /** Same idea as canManageAttachments(), scoped to Supporting Documents —
   *  that tab has no per-category split. Gated by canFullEdit only (NOT
   *  canEdit) — Data Science Programmer can still VIEW and DOWNLOAD
   *  existing supporting documents (those buttons have no permission
   *  gate at all, see the template), but cannot add, upload, delete, or
   *  submit/update this section. Public so the template can also hide/
   *  disable the relevant controls. */
  get canManageSupportingDocs(): boolean {
    return !this.isReadOnly && !this.isRecordLocked && this.canFullEdit;
  }

  /** Blocks a file action with a consistent toast and returns whether it
   *  was actually blocked, so callers can `if (this.blockIfCannotManage(...)) return;`. */
  private blockIfCannotManage(cat: AttachCategory): boolean {
    if (this.canManageAttachments(cat)) return false;
    this.toastr.error('You do not have permission to upload or delete files.', 'Access Denied');
    return true;
  }

  /** Same idea as blockIfCannotManage(), scoped to Supporting Documents —
   *  that tab has no per-category split (it's always gated by canEdit),
   *  so this just wraps the isReadOnly/canEdit check with the same toast. */
  private blockIfCannotManageDocs(): boolean {
    if (this.canManageSupportingDocs) return false;
    this.toastr.error('You do not have permission to upload or delete files.', 'Access Denied');
    return true;
  }

  /** Statuses the CURRENT role may move THIS concept's status to, given
   *  its current status — fetched from GET /api/allowed-statuses so the
   *  dropdown always matches the backend permission matrix in
   *  status_permissions.py. Populated by refreshAllowedStatuses(). */
  allowedStatusOptions: IdValueOption[] = [];

  /** True when the role has no permitted transitions away from the
   *  concept's current status (e.g. Operations/Viewer, or a role whose
   *  stage has already passed for this concept). Drives disabling the
   *  Development Status select even for roles that can otherwise edit. */
  statusLocked = false;

  /** Re-fetches allowedStatusOptions for the given current status and the
   *  logged-in user's role, then rebuilds the dropdown's option list so it
   *  always contains: (a) every status the role may transition TO, plus
   *  (b) the concept's own current status, so the field never renders
   *  blank/invalid for a value the role isn't allowed to move away from. */
  private lastKnownDevStatus = 'New';
  private refreshAllowedStatuses(currentStatus: string): void {
    this.lastKnownDevStatus = currentStatus || 'New';
    const userId = Number(sessionStorage.getItem('userId'));
    const roleName = sessionStorage.getItem('roleName') ?? '';

    if (this.isReadOnly || !userId) {
      this.allowedStatusOptions = this.developmentStatusOptions.filter(
        d => d.value === currentStatus
      );
      this.statusLocked = true;
      this.form.get('developmentStatus')?.disable({ emitEvent: false });
      return;
    }

    this.service.getAllowedStatuses(userId, currentStatus || 'New', roleName).subscribe({
      next: (res) => {
        // The API returns allowed_next_statuses as a list of
        // { id, name } objects (matching developmentStatusOptions'
        // { id, value } shape), not plain strings — pull out the
        // name so applyAllowedStatusFilter()'s
        // allowedStatuses.includes(d.value) string comparison actually
        // matches instead of silently comparing strings to objects and
        // never finding anything.
        const rawStatuses = res?.allowed_next_statuses ?? [];
        this.allowedStatuses = rawStatuses.map((s: any) =>
          typeof s === 'string' ? s : (s?.name ?? s?.value ?? '')
        ).filter((name: string) => !!name);

        // Backend is the source of truth for which statuses require
        // Estimated Volume / Estimated Dollars — record every status
        // this response tells us about (current + every reachable next
        // status) so estimatedFieldsRequired() / updateEstimatedFields
        // Validators() never have to guess.
        if (res?.current_status) {
          this.setStatusRequiresEstimates(res.current_status);
        }
        rawStatuses.forEach((s: any) => {
          if (s && typeof s === 'object') this.setStatusRequiresEstimates(s);
        });

        // patchForm() may have already set developmentStatus (and fired
        // updateEstimatedFieldsValidators against the '' default) before
        // this response landed — re-sync now that the map is populated.
        this.updateEstimatedFieldsValidators(this.form.get('developmentStatus')?.value);

        this.applyAllowedStatusFilter(currentStatus);
      },
      error: () => {
        this.allowedStatusOptions = this.developmentStatusOptions.filter(
          d => d.value === currentStatus
        );
        this.statusLocked = true;
        this.form.get('developmentStatus')?.disable({ emitEvent: false });
      }
    });
  }

  private applyAllowedStatusFilter(currentStatus: string): void {
    this.allowedStatusOptions = this.developmentStatusOptions.filter(
      d =>
        this.allowedStatuses.includes(d.value) ||
        d.value === currentStatus
    );

    this.statusLocked = this.allowedStatuses.length === 0;

    queueMicrotask(() => {
      const control = this.form.get('developmentStatus');

      // Production lock always wins over the normal allowed-transitions
      // logic below: only Manager keeps this field enabled, and every
      // other role stays locked out of it regardless of what the backend
      // says it could otherwise transition to.
      if (this.isRecordLocked) {
        if (this.isProductionLocked && this.isManager) {
          control?.enable({ emitEvent: false });
        } else {
          control?.disable({ emitEvent: false });
        }
        return;
      }

      if (this.statusLocked || this.isReadOnly) {
        control?.disable({ emitEvent: false });
      } else if (this.canEdit) {
        control?.enable({ emitEvent: false });
      }
    });
  }

  /** Disables form controls the current role is not allowed to edit.
 *  Called after patchForm() and BEFORE lockCoreFields() — see the order
 *  note in fetchAndApplyConcept() and the reset comment just inside this
 *  method for why that order matters. Re-called on resetToNewConcept()
 *  for the create flow (which has no lockCoreFields() call at all, since
 *  a blank concept has no id yet for those fields to lock against). */
  private applyRoleRestrictions(): void {
    // ── Reset to a clean, fully-enabled baseline first ─────────────────
    // This component instance (and its single reactive form) is reused
    // across concepts (see routeSub in ngOnInit) — patchValue() never
    // touches a control's enabled/disabled state, only its value. Without
    // this reset, switching FROM a fully-locked concept (Production /
    // Superseded / a read-only role's view) TO one a full-edit role
    // should be able to edit would leave every control still disabled:
    // e.g. every branch below either explicitly disables specific fields,
    // or — for canFullEdit roles — does nothing at all, silently relying
    // on the form already being enabled. That assumption breaks the
    // moment the concept just before it in the session was locked. Reset
    // here so every branch below starts from the same known state
    // regardless of whatever the previously-loaded concept left behind.
    // (lockCoreFields(), which intentionally stays locked forever once a
    // concept exists, deliberately runs AFTER this method — see
    // fetchAndApplyConcept() — so this reset can never undo it.)
    Object.keys(this.form.controls).forEach(key => {
      this.form.get(key)?.enable({ emitEvent: false });
    });
    this.form.get('confidenceScore')?.enable({ emitEvent: false });

    // ── Production lock ────────────────────────────────────────────────
    // Once this concept's Development Status is "Production", it's
    // finalized: every control on the form is disabled for every role.
    // Manager is the sole exception, and only for the Development Status
    // dropdown itself — re-enabled explicitly below — so a Manager can
    // still move the concept off Production or correct a mistake without
    // being able to touch anything else. This check runs before, and
    // takes priority over, every other branch in this method.
    if (this.isRecordLocked) {
      Object.keys(this.form.controls).forEach(key => {
        this.form.get(key)?.disable({ emitEvent: false });
      });
      this.form.get('confidenceScore')?.disable({ emitEvent: false });

      if (this.isProductionLocked && this.isManager) {
        this.form.get('developmentStatus')?.enable({ emitEvent: false });
      }
      return;
    }

    if (this.isReadOnly) {
      // Disable every control in the form for pure read-only roles.
      Object.keys(this.form.controls).forEach(key => {
        this.form.get(key)?.disable({ emitEvent: false });
      });
      // Also disable the confidence score nested group.
      this.form.get('confidenceScore')?.disable({ emitEvent: false });
      return;
    }

    if (this.canDSEdit) {
      // Data Science Programmer: on the Development tab, ONLY these fields
      // are editable —
      //   Internal Concept Description, Development Status, Priority,
      //   Development Notes, Estimated Volume, Estimated Dollars,
      //   Attachments, Data Science Programmer.
      // Ideation Requestor is shown disabled (read-only) as
      // "<name> - Ideation Requestor" — DS Programmer can see who
      // requested the concept but cannot reassign it.
      // Every other field (concept name, client/master/review/claim type,
      // halo number, confidence score, previous report id, QA/Production
      // schedule, and all Client Approval fields) is locked to read-only.
      const dsLockedFields = [
        'conceptName', 'clientName', 'masterConceptName',
        'reviewType', 'claimType', 'haloNumber',
        'previousReportId', 'qaSchedule', 'productionSchedule',
        'ideationRequestor',
        // Client Approval fields — DS Programmer has no approval access.
        'clientConceptName', 'clientConceptDescription',
        'clientApprovalStatus', 'submittedToClientOn', 'clientApprovalNotes',
        'clientEstimatedVolume', 'clientEstimatedDollars'
      ];
      dsLockedFields.forEach(f => this.form.get(f)?.disable({ emitEvent: false }));

      // These are the ONLY fields DS Programmer CAN edit — ensure they're
      // enabled (the reset at the top of this method already did this,
      // but stating it explicitly keeps this branch self-contained).
      // Development Notes and Attachments are not reactive-form controls,
      // so they're gated separately (see canManageAttachments() and the
      // Development Notes input, which is left unrestricted for any role
      // that isn't pure read-only).
      const dsEditableFields = [
        'developmentStatus', 'priority',
        'Internalconceptdescription', 'estimatedVolume', 'estimatedDollars',
        'dataScienceProgrammer'
      ];
      dsEditableFields.forEach(f => this.form.get(f)?.enable({ emitEvent: false }));

      // Confidence Score is NOT in the DS Programmer's allowed field list —
      // lock the nested form group (this also backs up the [disabled]
      // binding added to the toggle buttons in the template, since those
      // buttons call setConfidenceScore() directly and aren't otherwise
      // blocked by form.disable()).
      this.form.get('confidenceScore')?.disable({ emitEvent: false });
    }
    // canFullEdit roles: no extra restrictions beyond lockCoreFields().
  }


  private originalFieldValues: Record<string, any> = {};
  private readonly watchedFields = [
    'conceptName',
    'developmentStatus',
    'priority',
    'estimatedVolume',
    'estimatedDollars',
    'confidenceScore',   // nested group — handled specially below
    'qaSchedule',
    'productionSchedule',
    'ideationRequestor',
    'dataScienceProgrammer',
    'haloNumber',
    'previousReportId',
    'Internalconceptdescription',
  ];

  // Attachment categories that live on the Concept Development tab —
  // 'approval' belongs to Client Approval and isn't part of this check.
  private readonly trackedAttachmentCategories: AttachCategory[] = ['specs', 'table', 'other'];

  // Client Approval tab's own set of tracked fields, checked via each
  // control's .dirty flag (rather than a value snapshot like watchedFields)
  // since patchClientApproval() already carefully maintains dirty state for
  // exactly this purpose (see its isConceptSwitch branches). Deliberately
  // excludes estimatedVolume/estimatedDollars — those are shared with the
  // Development tab's watchedFields, so a dirty flag left over from an
  // unrelated Development-tab edit would incorrectly make this tab think
  // it has its own unsaved change.
  private readonly approvalFields = [
    'clientConceptName',
    'clientConceptDescription',
    'clientApprovalStatus',
    'submittedToClientOn',
    'clientApprovalNotes',
    'clientEstimatedVolume',
    'clientEstimatedDollars',
  ];

  /** True if any Client Approval field has been edited since the last load
   *  or successful submit — mirrors getChangedWatchedFields() but keyed off
   *  Angular's own dirty flag instead of a value snapshot. Used by
   *  onApprovalSubmit()'s "did anything change" guard. */
  private getApprovalFieldsChanged(): boolean {
    return this.approvalFields.some(field => !!this.form.get(field)?.dirty);
  }

  // Snapshot of AttachmentIds present right after load/refresh, per
  // category — used the same way originalFieldValues is: anything added
  // (no attachmentId yet) or missing from this snapshot counts as a change.
  private originalAttachmentIds: Record<AttachCategory, Set<number>> = {
    specs: new Set(), table: new Set(), other: new Set(), approval: new Set()
  };

  private snapshotAttachmentIds(): void {
    this.originalAttachmentIds = { specs: new Set(), table: new Set(), other: new Set(), approval: new Set() };
    for (const cat of this.trackedAttachmentCategories) {
      this.attachments[cat].forEach(f => {
        if (f.attachmentId) this.originalAttachmentIds[cat].add(f.attachmentId);
      });
    }
  }

  /** True if any attachment in any of the given categories has been added
   *  or removed since the last snapshot — shared by getAttachmentsChanged()
   *  (Development tab: specs/table/other) and getApprovalAttachmentsChanged()
   *  (Client Approval tab: approval), so both tabs' "did anything actually
   *  change" checks compare against the same snapshot mechanism. */
  private getAttachmentsChangedForCategories(categories: AttachCategory[]): boolean {
    return categories.some(cat => {
      const current = this.attachments[cat];
      const hasNewFile = current.some(f => !f.attachmentId);
      if (hasNewFile) return true;

      const currentIds = new Set(current.filter(f => f.attachmentId).map(f => f.attachmentId!));
      const original = this.originalAttachmentIds[cat];
      if (currentIds.size !== original.size) return true;
      for (const id of original) {
        if (!currentIds.has(id)) return true;
      }
      return false;
    });
  }

  /** True if any specs/table/other attachment has been added or removed
   *  since the last snapshot — mirrors getChangedWatchedFields() but for
   *  attachments instead of form fields. */
  private getAttachmentsChanged(): boolean {
    return this.getAttachmentsChangedForCategories(this.trackedAttachmentCategories);
  }

  /** Client Approval tab's own version of getAttachmentsChanged() — checks
   *  only the 'approval' attachment category, which getAttachmentsChanged()
   *  deliberately excludes (see trackedAttachmentCategories). Used by
   *  onApprovalSubmit()'s "did anything change" guard, same purpose as the
   *  Development tab's check at the top of onSubmit(). */
  private getApprovalAttachmentsChanged(): boolean {
    return this.getAttachmentsChangedForCategories(['approval']);
  }

  get reversedDocs(): SupportingDoc[] {
    return [...this.supportingDocs].reverse();
  }

  /** Drives the header status badge — falls back to "New" before a
   *  development status has been chosen (e.g. a brand-new concept). */
  get developmentStatusLabel(): string {
    return this.form.get('developmentStatus')?.value || 'New';
  }

  /** CSS-safe slug of the label above, e.g. "In Progress" -> "in-progress",
   *  used as `status-{slug}` so each development status can get its own
   *  badge color (see .status-badge.status-* rules in the stylesheet). */
  get developmentStatusSlug(): string {
    return this.developmentStatusLabel
      .toLowerCase()
      .trim()
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/(^-|-$)/g, '');
  }

  get filteredCandidates(): OwnerCandidate[] {
    const q = this.ownerSearch.toLowerCase();
    const addedIds = new Set(this.owners.map(o => o.id));
    return this.ownerCandidates.filter(
      c => !addedIds.has(c.id) && (c.name.toLowerCase().includes(q) || c.department.toLowerCase().includes(q))
    );
  }

  // ── Development Notes ─────────────────────────────────────────────────
  newNoteText = '';
  devNotes: DevNote[] = [];

  // Drives the red outline on the note input — set true when a submit is
  // blocked because tracked fields/attachments changed but no note was
  // entered, cleared as soon as the user starts typing a note.
  noteInputInvalid = false;

  // Collapsible Development Notes card — collapsed by default so it
  // doesn't take up vertical space until the user actually wants to
  // read or add notes. Forced open below (see onSubmit()'s
  // noteInputInvalid block) whenever the user needs to see something
  // inside it, e.g. the "Development Note required" validation hint.
  notesExpanded = false;

  toggleNotesExpanded(): void {
    this.notesExpanded = !this.notesExpanded;
  }

  /** These two flags exist because Estimated Volume and Estimated Dollars
   *  are the SAME FormControl instances, reused on both the Concept
   *  Development tab and the Client Approval tab (see the requiredFields
   *  checks in onSubmit() and onApprovalSubmit()). Angular's control-level
   *  `.touched` state has no concept of "which tab touched it" — so
   *  calling markAsTouched() on a failed Concept Development submit also
   *  flips `.touched` for the exact same control instance the Client
   *  Approval tab is looking at, and any template binding keyed off
   *  `form.get('estimatedVolume').touched` lights up on BOTH tabs even
   *  though Client Approval was never opened or submitted.
   *
   *  Use these instead of `.touched` in the template to decide whether to
   *  SHOW an error for estimatedVolume/estimatedDollars (and any other
   *  field shared between the two tabs) — e.g.:
   *    *ngIf="developmentSubmitAttempted && form.get('estimatedVolume').invalid"
   *  on the Concept Development tab, and
   *    *ngIf="approvalSubmitAttempted && form.get('estimatedVolume').invalid"
   *  on the Client Approval tab. Each tab's own Submit click flips only
   *  its own flag (see onSubmit() / onApprovalSubmit() below), so a failed
   *  submit on one tab can never surface an error message on the other. */
  developmentSubmitAttempted = false;
  approvalSubmitAttempted = false;

  /** Bound to the note input's (input) event — clears the red-outline
   *  state the moment the user starts addressing it. */
  onNoteInputChange(): void {
    if (this.noteInputInvalid) this.noteInputInvalid = false;
  }

  /** Grows the Development Notes textarea to fit its content (instead of
   *  scrolling the text sideways like a single-line input), capped by
   *  the .note-input max-height in CSS so a very long note scrolls
   *  internally rather than growing forever. Also toggles the
   *  "multiline" class on the surrounding .note-input-wrap so the
   *  pill-shaped border relaxes into a standard rounded rect once the
   *  note wraps past one line. */
  autoResizeNoteInput(el: HTMLTextAreaElement): void {
    el.style.height = 'auto';
    el.style.height = `${el.scrollHeight}px`;
    el.parentElement?.classList.toggle(
      'note-input-multiline',
      el.scrollHeight > el.clientHeight || el.value.includes('\n') || el.scrollHeight > 40
    );
  }

  /** Resets the note textarea back to its single-line pill shape —
   *  called wherever newNoteText is cleared programmatically (after a
   *  successful save, cancel, etc.) since clearing the model alone
   *  doesn't shrink an element whose height was set inline by
   *  autoResizeNoteInput(). */
  private resetNoteInputHeight(): void {
    setTimeout(() => {
      const el = document.querySelector<HTMLTextAreaElement>('.note-input');
      if (el) {
        el.style.height = 'auto';
        el.parentElement?.classList.remove('note-input-multiline');
      }
    });
  }

  // Logged-in user's display name + initial — used for the "self" avatar
  // on the note composer row and for any note added in this session.
  // NOTE: adjust the sessionStorage key below if the app stores the
  // logged-in user's name under a different key elsewhere.
  currentUserName = sessionStorage.getItem('userName') || 'You';
  currentUserInitial = this.getInitial(this.currentUserName);

  private getInitial(name: string): string {
    return name?.trim()?.charAt(0)?.toUpperCase() || '?';
  }

  // ── Attachments ───────────────────────────────────────────────────────
  attachments: Record<AttachCategory, AttachFile[]> = {
    specs: [],
    table: [],
    other: [],
    approval: []
  };

  /** Next free DocIndex to hand out to a brand-new Supporting Document
   *  slot. Ratchets forward only — never reused, and never derived from
   *  `supportingDocs.length` (which shrinks when a doc is deleted). This
   *  guarantees a new slot's docIndex can never collide with any slot
   *  that currently exists OR previously existed for this concept. Reset
   *  (to 0) only when starting a brand-new concept (resetToNewConcept)
   *  or bumped up past every restored DocIndex when loading an existing
   *  one (see patchSupportingDocs). */
  private nextDocIndex = 0;

  private allocateDocIndex(): number {
    return this.nextDocIndex++;
  }

  /** Builds a brand-new, empty Supporting Document slot with a freshly
   *  allocated, stable docIndex. Use this instead of a raw object literal
   *  anywhere a NEW slot is created — never assign docIndex from the
   *  slot's position in the array. */
  private blankDoc(): SupportingDoc {
    return {
      name: '', sourceurl: '', pdfLocation: '', uploadProgress: 0, file: null,
      docIndex: this.allocateDocIndex()
    };
  }

  // ── Supporting Documents ──────────────────────────────────────────────
  supportingDocs: SupportingDoc[] = [
    this.blankDoc(),
    this.blankDoc(),
    this.blankDoc()
  ];

  /** True while any specs/table/other attachment's progress bar hasn't
   *  reached 100% yet. Drives [disabled] on Save as Draft / Create
   *  Concept / Update so a click can't fire while a file is still
   *  "uploading" (see simulateUploadProgress). */
  get isAttachmentUploading(): boolean {
    const categories: AttachCategory[] = ['specs', 'table', 'other'];
    return categories.some(cat => this.attachments[cat].some(f => f.progress < 100));
  }

  /** Same idea for the Supporting Documents tab — any doc that has a file
   *  attached but whose progress bar hasn't reached 100% yet blocks Submit. */
  get isSupportingDocUploading(): boolean {
    return this.supportingDocs.some(d => d.file && d.uploadProgress < 100);
  }

  /** Same idea, scoped to the Client Approval tab's own attachment
   *  category — blocks the Approval Submit button while a file there is
   *  still mid-upload, without affecting the Development tab's button. */
  get isApprovalAttachmentUploading(): boolean {
    return this.attachments.approval.some(f => f.progress < 100);
  }

  // ── Generic delete/clear confirmation popup ─────────────────────────────
  // Shared "Are you sure?" gate for every destructive delete/remove/clear
  // action on this page. Opening the popup never performs the action —
  // the caller hands over the action as a callback, and it only runs if
  // the user confirms (confirmPendingAction()). Cancelling, closing the
  // popup, or clicking the backdrop all leave the underlying data
  // untouched.
  showConfirmModal = false;
  confirmModalMessage = '';
  confirmModalConfirmLabel = 'Delete';
  private pendingConfirmAction: (() => void) | null = null;

  private openConfirmModal(
    message: string,
    action: () => void,
    confirmLabel: string = 'Delete'
  ): void {
    this.confirmModalMessage = message;
    this.confirmModalConfirmLabel = confirmLabel;
    this.pendingConfirmAction = action;
    this.showConfirmModal = true;
  }

  confirmPendingAction(): void {
    const action = this.pendingConfirmAction;
    this.showConfirmModal = false;
    this.pendingConfirmAction = null;
    action?.();
  }

  cancelConfirmModal(): void {
    this.showConfirmModal = false;
    this.pendingConfirmAction = null;
  }

  /** Confirmation wrapper for removeAttachment() — used by the SPECS,
   *  TABLE, OTHER, and Client Approval attachment cards. */
  confirmRemoveAttachment(cat: AttachCategory, f: AttachFile): void {
    this.openConfirmModal(
      'This will permanently delete this file.',
      () => { void this.removeAttachment(cat, f); },
      'Delete'
    );
  }

  /** Confirmation wrapper for removeSupportingDoc() — removes the whole
   *  Supporting Document slot. */
  confirmRemoveSupportingDoc(index: number): void {
    this.openConfirmModal(
      'This will permanently delete this supporting document.',
      () => { void this.removeSupportingDoc(index); },
      'Delete'
    );
  }

  /** Confirmation wrapper for clearDocFile() — resets a Supporting
   *  Document slot back to blank without removing the slot itself. */
  confirmClearDocFile(event: Event, index: number): void {
    event.stopPropagation();
    this.openConfirmModal(
      'This will clear the file, name, and links from this document slot.',
      () => { this.clearDocFile(event, index); },
      'Clear'
    );
  }

  /** Confirmation wrapper for removeOwner(). */
  confirmRemoveOwner(id: string): void {
    this.openConfirmModal(
      'This will remove this owner from the concept.',
      () => { this.removeOwner(id); },
      'Remove'
    );
  }

  // ── SPECS replace-confirmation ──────────────────────────────────────────
  // SPECS only ever holds a single file. Picking a new one while an
  // existing SPECS file is already attached doesn't swap it silently —
  // the user has to confirm deactivating the current file first. These
  // hold the file/entry pending that confirmation.
  showSpecsReplaceModal = false;
  private pendingSpecsFile: File | null = null;
  private pendingSpecsExisting: AttachFile | null = null;

  /** The persisted SPECS file (if any) that was deactivated by a confirmed
   *  replace but is NOT yet deleted on the backend. New attachments aren't
   *  actually uploaded until submitConcept() succeeds (see filesToUpload
   *  there), so deleting the old file any earlier — e.g. right when the
   *  user confirms the popup — would leave SPECS with zero backend files
   *  if they refresh or navigate away before hitting Update/Submit. The
   *  real deleteattachment call only fires from submitConcept() once the
   *  replacement has actually been saved; see the cleanup block there. */
  private pendingSpecsDeletion: AttachFile | null = null;

  get pendingSpecsFileName(): string {
    return this.pendingSpecsFile?.name ?? '';
  }

  get pendingSpecsExistingName(): string {
    return this.pendingSpecsExisting?.name ?? '';
  }

  // ── Superseded → new development cycle switch-confirmation ─────────────
  // A Manager moving this concept's Development Status from Production to
  // Superseded locks the current record fully read-only (see
  // isProductionLocked/isReadOnly) and the backend spins up a brand-new
  // development cycle in the same response (new_concept_id). Rather than
  // silently leaving the Manager stranded on the now-locked record, or
  // silently jumping them away from the save they just made, surface a
  // confirmation popup: they stay on the (now read-only) superseded
  // concept until they explicitly choose to switch.
  showSupersededSwitchModal = false;
  private pendingSupersededConceptId: string | null = null;

  get pendingSupersededConceptLabel(): string {
    return this.pendingSupersededConceptId ?? '';
  }

  // ── Upload Modal ──────────────────────────────────────────────────────
  showUploadModal = false;
  modalSelectedIndex: number | null = null;

  // ── Doc Viewer (Word/Excel/PDF popup) ──────────────────────────────────
  docViewerVisible = false;
  docViewerFileName = '';
  docViewerFileExt = '';
  docViewerLoading = false;
  docWordHtml: SafeHtml = '';
  docSheets: SheetData[] = [];
  docActiveSheet = 0;
  docPdfUrl: SafeResourceUrl | null = null;
  private docPdfObjectUrl: string | null = null;   // raw URL, kept to revoke on close


  private allowedStatuses: string[] = [];

  // ── Submit / Upload state ─────────────────────────────────────────────
  // `loading` drives the Concept Development tab's Submit/Create/Update
  // button (and is reused as-is by the Client Approval / Supporting
  // Documents tabs' own Submit buttons, each on their own tab). Save as
  // Draft gets its OWN flag, `savingDraft`, even though both actions run
  // through the same submitConcept() pipeline — sharing `loading` between
  // the two meant clicking Save as Draft also flipped the Submit button
  // into its disabled/"Uploading..." state for the duration of the draft
  // save, which read as the Submit button "blinking" even though the user
  // never touched it.
  loading = false;
  savingDraft = false;

  constructor(
    private fb: FormBuilder,
    private cdr: ChangeDetectorRef,
    private service: Service,
    private toastr: ToastrService,
    private sanitizer: DomSanitizer,
    private route: ActivatedRoute,
    private router: Router
  ) {
    // Captured here, not in ngOnInit: navigating from /concept-create/:id
    // to /concept-create is a DIFFERENT route config, not just a changed
    // :id param, so Angular's default reuse strategy tears down this
    // whole component and constructs a brand-new instance for it rather
    // than reusing this one — that's exactly why patching `this.form`
    // synchronously after calling router.navigate() (the old approach)
    // silently did nothing: it patched the outgoing instance, which was
    // already on its way out, while an entirely new instance (with its
    // own fresh `this.form`) is what actually ends up on screen.
    // router.getCurrentNavigation() is only non-null while a navigation
    // this component is being constructed FOR is still in flight, so
    // reading it here — in the constructor of the freshly-created
    // instance — is the one reliable place to pick up state passed via
    // router.navigate(..., { state: {...} }) regardless of whether the
    // instance was reused or recreated. See applyCopiedConcept() for
    // where this is set, and the routeSub below for where it's consumed.
    const nav = this.router.getCurrentNavigation();
    this.pendingCopiedMetadata = (nav?.extras?.state as any)?.copiedConceptMetadata ?? null;
  }

  /** Metadata from a just-completed Copy Concept action, staged here (via
   *  router.navigate state — see the constructor) so it survives the
   *  /concept-create/:id -> /concept-create navigation even if that
   *  navigation recreates this component. Consumed and cleared by the
   *  routeSub below the first time it lands on the blank creation route. */
  private pendingCopiedMetadata: any = null;

  private routeSub?: Subscription;

  // ── Lifecycle ─────────────────────────────────────────────────────────
  ngOnInit(): void {
    this.buildForm();
    this.loadLatestUpdates();
    this.loadMasterData();
    if (this.activeTab === 'development') {
      // this.loadUserFiles();
    }

    // Re-run whenever the :id route param changes, not just on first load.
    // (Angular reuses this component instance when navigating between
    // /concept-create/:id routes, so ngOnInit itself only fires once —
    // reading route.snapshot here would go stale on subsequent navigations.)
    this.routeSub = this.route.paramMap.subscribe(params => {
      const routeId = params.get('id');
      const queryId = this.route.snapshot.queryParamMap.get('id');
      const id = routeId ?? queryId;

      // Angular REUSES this component instance across /concept-create/:id
      // navigations (that's the whole reason routeSub exists) — so any
      // per-concept transient UI state has to be explicitly cleared here,
      // or it leaks from the concept you just left into the one you're
      // navigating to. The draft-lock banner is exactly that: it's only
      // ever set true in flashDraftLockBanner() and only ever cleared by
      // onTabChange()/dismissDraftLockBanner() — neither of which fires
      // just because you selected a different concept from the Latest
      // Updates list. Without this, switching away from a draft concept
      // (after triggering the banner) to an already-submitted concept
      // keeps showing "This concept is still a draft…" about the NEW
      // concept, even though it isn't one.
      this.dismissDraftLockBanner();

      if (id) {
        this.isEditMode = true;
        this.conceptId = id;
        this.loadConcept(id);
      } else if (!this.canCreateConcept) {
        // Roles without create rights (Data Science Programmer, Viewer,
        // Operations) have no business on the blank "new concept" form —
        // they can't submit it, and the "+ Add New Concept" button that
        // would normally get them here is already disabled for them on
        // the Dashboard. This only matters for someone landing on
        // /concept-create directly (typed URL, stale bookmark, browser
        // back/forward), so send them to the most recently updated
        // concept instead, same as clicking the top card in Latest
        // Updates.
        this.redirectToMostRecentConcept();
      } else {
        this.isEditMode = false;
        this.resetToNewConcept();
        // Apply a just-copied concept's metadata now that the blank form
        // it belongs on has actually been (re)built above — see the
        // constructor for why this can't just be patched directly by
        // whatever triggered this navigation.
        if (this.pendingCopiedMetadata) {
          const metadata = this.pendingCopiedMetadata;
          this.pendingCopiedMetadata = null;
          this.applyCopiedMetadataToForm(metadata);
        }
      }
    });
  }

  /** Sends a role that can't create concepts to the most recently
   *  updated concept, in place of the blank creation form. Fetches its
   *  own copy of /api/latest-updates rather than reading this.latestConcepts
   *  — ngOnInit's loadLatestUpdates() call races this routeSub callback
   *  and usually hasn't resolved yet, so this.latestConcepts is still []
   *  at this point most of the time. Sorted identically to
   *  loadLatestUpdates() (most recently updated/created first) so "first
   *  in the list" here matches what the Latest Updates panel itself would
   *  show as its top card. */
  private redirectToMostRecentConcept(): void {
    this.service.getLatestUpdates().subscribe({
      next: (res) => {
        const concepts: any[] = res?.data ?? [];
        const mostRecent = concepts
          .slice()
          .sort((a, b) => {
            const bTime = new Date(b.UpdatedDate ?? b.CreatedDate).getTime();
            const aTime = new Date(a.UpdatedDate ?? a.CreatedDate).getTime();
            return bTime - aTime;
          })[0];

        // MUST route on the stable anchor (ConceptId), never the
        // version/display id (CurrentConceptId) — see the routing note
        // on onSelectLatestConcept().
        const targetId = mostRecent?.ConceptId;

        if (targetId) {
          this.router.navigate(['/concept-create', targetId], { replaceUrl: true });
        } else {
          // No concepts exist anywhere yet — nothing to redirect to.
          // Fall back to the blank form rather than a dead end; every
          // field on it is already read-only/disabled for this role.
          this.isEditMode = false;
          this.resetToNewConcept();
        }
      },
      error: (err) => {
        console.error('Failed to load latest updates for redirect:', err);
        this.isEditMode = false;
        this.resetToNewConcept();
      }
    });
  }
  ngOnDestroy(): void {
    this.routeSub?.unsubscribe();
    if (this.draftLockBannerTimer) {
      clearTimeout(this.draftLockBannerTimer);
    }
  }

  // ── Load existing concept ─────────────────────────────────────────────
  // Used for the initial route-driven load — shows the page-level loading
  // state and always lands on the Development tab.
  private loadConcept(id: string): void {
    this.pageLoading = true;
    this.activeTab = 'development';
    this.developmentSubmitAttempted = false;
    this.approvalSubmitAttempted = false;
    this.noteInputInvalid = false;
    this.newNoteText = '';
    this.resetNoteInputHeight();
    // Genuine concept switch — this component instance is reused across
    // concepts (see routeSub in ngOnInit / fetchAndApplyConcept), so any
    // per-concept UI toggle left open on the PREVIOUS concept must not
    // carry over onto this one. Same reasoning as resetToNewConcept().
    this.notesExpanded = false;
    this.showFullActivityHistory = false;
    this.selectedActivity = null;
    this.showActivityLegend = false;
    this.activityDotLegend = [];
    this.fetchAndApplyConcept(id, () => {
      this.pageLoading = false;
      this.scrollActiveConceptIntoView();
      this.scrollTabContentToTop();
    }, 'Failed to load concept', /* isConceptSwitch */ true);
  }

  /** Soft-reload: re-fetches this concept from the server and re-patches
   *  every tab in place right after a successful save/update, so the page
   *  reflects exactly what was persisted (computed dates, server-side
   *  defaults, notes, attachment state, etc.) instead of relying on the
   *  optimistic local mutations made right after the save call. No
   *  page-level loading flag and no tab switch — so it never disrupts
   *  whichever tab the user is currently on — and no full browser page
   *  reload either.
   *
   *  Called for the SAME concept that's already loaded (e.g. right after
   *  a Supporting Documents or Client Approval submit succeeds), so this
   *  is a same-concept soft refresh, not a concept switch — see
   *  isConceptSwitch on fetchAndApplyConcept() / patchClientApproval(). */
  /** @param resnapshotDevelopmentFields Only true when this refresh follows
   *  the Concept Development tab's OWN successful Update/Create — see the
   *  note on fetchAndApplyConcept() below for why every other caller must
   *  leave this false. */
  private refreshConceptData(id: string, resnapshotDevelopmentFields: boolean = false): void {
    if (!id) return;
    this.fetchAndApplyConcept(
      id, undefined, 'Failed to refresh concept',
      /* isConceptSwitch */ false, resnapshotDevelopmentFields
    );
  }

  /** Shared fetch + patch logic behind loadConcept() and
   *  refreshConceptData() — keeps both call sites in sync with whatever
   *  the form actually needs patched.
   *
   *  isConceptSwitch distinguishes the two very different reasons this
   *  can run:
   *   - true  (loadConcept / initial route load): a genuine navigation to
   *     a different concept (or the first load of this one). Every tab,
   *     including Client Approval, must be fully reset to exactly what
   *     the server has for THIS concept — this component instance is
   *     reused across concepts (see routeSub in ngOnInit), so stale data
   *     from whatever was open before must not leak in.
   *   - false (refreshConceptData): a same-concept soft refresh, fired
   *     right after some OTHER tab's submit succeeds (Supporting
   *     Documents, Client Approval itself, etc). The user may currently
   *     have unsaved, not-yet-submitted edits sitting in the Client
   *     Approval controls — the server hasn't seen them yet, so
   *     blindly re-patching from res.client_approvals would wipe out
   *     in-progress work and, because those controls are already
   *     touched/dirty, immediately flip them to red "required" errors.
   *     See patchClientApproval().
   *
   *  resnapshotDevelopmentFields guards the SAME kind of problem for the
   *  Concept Development tab's own "did a watched field change since
   *  the last save?" tracking (see watchedFields / originalFieldValues /
   *  getChangedWatchedFields(), used by onSubmit() to require a
   *  Development Note before Updating). estimatedVolume and
   *  estimatedDollars are shared, editable fields on BOTH the Concept
   *  Development and Client Approval tabs. Submitting Client Approval
   *  (or Supporting Documents) triggers this same soft refresh, and
   *  patchForm() always re-patches those shared fields from whatever the
   *  server now has. Re-baselining (snapshotFieldValues()) on THAT
   *  refresh would silently treat an Estimated Volume/Dollars edit that
   *  was only ever persisted via Client Approval as if it had been
   *  cleanly saved from the Development tab too — erasing the "you
   *  changed this, add a note" state even though no Development Note
   *  was ever provided, and even though the note input is still showing
   *  its red "required" outline. So only the Development tab's own
   *  submitConcept() success (the one call site that actually satisfied
   *  the note requirement, or legitimately didn't need to) may pass
   *  true here; every other soft refresh must leave the existing
   *  baseline alone so a still-unexplained change keeps being reported
   *  as changed until the user actually adds a note and updates from
   *  the Development tab itself. */
  private fetchAndApplyConcept(
    id: string,
    onDone?: () => void,
    errorMessage: string = 'Failed to load concept',
    isConceptSwitch: boolean = true,
    resnapshotDevelopmentFields: boolean = isConceptSwitch
  ): void {
    this.service.getConcept(id).subscribe({
      next: (res) => {
        // Stale-response guard: this component instance is reused across
        // concepts (see routeSub in ngOnInit), and this fetch is not
        // cancelled if the user navigates away before it resolves. That
        // race is real, not hypothetical — e.g. submitConcept()'s own
        // soft-reload of THIS (just-superseded) concept is still in
        // flight, unawaited, when confirmSwitchToSupersededConcept()
        // immediately navigates to the newly-created concept and fires a
        // second, independent fetch for it. If the superseded concept's
        // response then lands after the new concept's, it would silently
        // overwrite the page with the old, now-read-only record — which
        // is exactly the "switches in read-only, but a manual refresh
        // fixes it" bug this guard prevents. This.conceptId is updated
        // synchronously by routeSub BEFORE loadConcept()/this fetch ever
        // fires (and synchronously by submitConcept() for the
        // same-concept refresh case), so if it no longer matches the id
        // THIS response is for, something newer has already superseded
        // this request — discard it instead of patching stale data in.
        if (id !== this.conceptId) {
          return;
        }

        const c = res.concept ?? {};
        const files = res.active_files ?? [];

        // Genuine concept switch: this component instance is reused across
        // concepts (see routeSub in ngOnInit), and patchValue() below never
        // touches Angular's touched/dirty flags — it only changes values.
        // Left alone, any field a user had blurred/edited on the PREVIOUS
        // concept (or on the blank "new concept" form) would stay marked
        // touched/dirty here, so this brand-new concept's own untouched
        // fields would immediately render red "required" errors even
        // though nobody has interacted with them yet. Clear that state
        // before patching in this concept's data. (On a same-concept soft
        // refresh — isConceptSwitch === false — this must NOT run: that's
        // exactly the in-progress touched/dirty state we need to preserve,
        // see patchClientApproval().)
        if (isConceptSwitch) {
          this.form.markAsUntouched();
          this.form.markAsPristine();
        }

        this.patchForm(c);
        this.snapshotPersistedFieldValues();
        // Refresh the Activity History card every time concept data is
        // (re)loaded — both on a genuine concept switch AND on the
        // soft-refresh path (e.g. right after a Client Approval or
        // Supporting Documents submit), since those actions generate new
        // audit rows this same request needs to reflect.
        this.loadActivityHistory(id);
        // this.cdr.detectChanges();
        if (resnapshotDevelopmentFields) {
          this.snapshotFieldValues();
          // conceptName is intentionally excluded from watchedFields (see
          // its declaration — editing it alone doesn't require a
          // Development Note), but onSubmit()'s separate "did anything
          // change at all" check does look at its dirty flag. Angular's
          // dirty flag isn't cleared by patchValue()/patchForm() above, so
          // without this it would stay true forever after the first edit,
          // even once that edit's been legitimately saved — making that
          // check think there's still an unsaved change on every future
          // Update click.
          this.form.get('conceptName')?.markAsPristine();
        }
        // Order matters: applyRoleRestrictions() now resets every control
        // to enabled before applying its own restrictions (see its
        // comment), so it must run FIRST — otherwise it would immediately
        // undo lockCoreFields()'s always-on lock (clientName,
        // masterConceptName, reviewType, claimType, ideationRequestor)
        // the instant this callback runs.
        this.applyRoleRestrictions();
        this.lockCoreFields();
        this.refreshAllowedStatuses(c.DevelopmentStatus ?? 'New');
        this.patchMeta(c);
        this.patchAttachments(files);                          // Attachments tab (SPECS / TABLE / OTHER)
        this.snapshotAttachmentIds();
        this.patchSupportingDocs(files);                       // Supporting Documents tab
        this.patchDevNotes(res.development_notes ?? []);       // Development Notes
        this.patchClientApproval(
          res.client_approvals ?? [],
          isConceptSwitch,
          { volume: c.EstimatedVolume ?? '', dollars: c.EstimatedDollars ?? '' }
        );  // Client Approval tab
        if (isConceptSwitch) {
          this.ensureClientConceptName();
        }
        // this.cdr.detectChanges();
        onDone?.();
      },
      error: (err) => {
        // Same staleness guard as the success branch above — an error
        // from a superseded, no-longer-relevant fetch (e.g. the old
        // concept's soft-reload failing after the user already switched
        // away) shouldn't surface an error toast for a request nobody's
        // waiting on anymore, and definitely shouldn't call the NEW
        // request's onDone (e.g. clearing pageLoading) out of turn.
        if (id !== this.conceptId) {
          return;
        }
        this.toastr.error(errorMessage, 'Error');
        console.error(err);
        onDone?.();

      }
    });
  }
  private snapshotFieldValues(): void {
    this.originalFieldValues = {};
    for (const field of this.watchedFields) {
      if (field === 'confidenceScore') {
        this.originalFieldValues[field] = this.form.get('confidenceScore.value')?.value ?? '';
      } else {
        this.originalFieldValues[field] = this.form.get(field)?.value ?? '';
      }
    }
  }

  private getChangedWatchedFields(): string[] {
    return this.watchedFields.filter(field => {
      const original = this.originalFieldValues[field] ?? '';
      const current = field === 'confidenceScore'
        ? (this.form.get('confidenceScore.value')?.value ?? '')
        : (this.form.get(field)?.value ?? '');
      return String(current).trim() !== String(original).trim();
    });
  }
  private lastPersistedFieldValues: Record<string, any> = {};

  private snapshotPersistedFieldValues(): void {
    this.lastPersistedFieldValues = {};
    for (const field of this.watchedFields) {
      if (field === 'confidenceScore') {
        this.lastPersistedFieldValues[field] = this.form.get('confidenceScore.value')?.value ?? '';
      } else {
        this.lastPersistedFieldValues[field] = this.form.get(field)?.value ?? '';
      }
    }
  }

  private getFieldsDifferFromPersisted(): string[] {
    return this.watchedFields.filter(field => {
      const persisted = this.lastPersistedFieldValues[field] ?? '';
      const current = field === 'confidenceScore'
        ? (this.form.get('confidenceScore.value')?.value ?? '')
        : (this.form.get(field)?.value ?? '');
      return String(current).trim() !== String(persisted).trim();
    });
  }

  private patchForm(c: any): void {
    // Persisted value as of this load — this is what the Client Approval
    // tab's clientConceptName is allowed to mirror, not whatever's
    // currently (possibly unsaved) sitting in the conceptName control.
    this.savedConceptName = (c.ConceptName ?? '').toString();

    // Snapshot the concept's own persisted Development Status — drives
    // isProductionLocked (see its getter). Must be set before
    // lockCoreFields()/applyRoleRestrictions() run in fetchAndApplyConcept().
    this.currentConceptStatus = c.DevelopmentStatus ?? '';
    if (this.currentConceptStatus === 'Production'|| this.currentConceptStatus === 'Superseded') {
      this.flashProductionLockBanner();
    } else {
      this.dismissProductionLockBanner();
    }

    this.form.patchValue({
      conceptName: c.ConceptName ?? '',
      clientName: c.ClientId ?? '',
      masterConceptName: c.MasterConceptId ?? '',
      reviewType: c.ReviewType ?? '',
      claimType: c.ClaimType ?? '',
      Internalconceptdescription: c.InternalConceptDescription ?? '',
      developmentStatus: c.DevelopmentStatus ?? '',
      priority: c.Priority ?? '',
      haloNumber: c.HaloNumber ?? '',
      estimatedVolume: c.EstimatedVolume ?? '',
      estimatedDollars: c.EstimatedDollars ?? '',
      previousReportId: c.PreviousReportId ?? '',
      qaSchedule: c.QASchedule ? c.QASchedule.split('T')[0] : '',
      productionSchedule: c.ProductionSchedule ? c.ProductionSchedule.split('T')[0] : '',
      ideationRequestor: c.IdeationRequestorId ?? '',
      dataScienceProgrammer: c.DataScienceProgrammerId ?? '',

      // // Client Approval (not present in this payload — keep as-is/blank)
      // clientConceptName:        c.ClientConceptName        ?? '',
      // clientConceptDescription: c.ClientConceptDescription ?? '',
      // clientApprovalStatus:     c.ClientApprovalStatus     ?? '',
      // submittedToClientOn: c.SubmittedToClientOn ? c.SubmittedToClientOn.split('T')[0] : '',
      // clientApprovalNotes:      c.ClientApprovalNotes      ?? '',
    });

    if (c.ConfidenceScore) {
      this.form.get('confidenceScore.value')?.setValue(c.ConfidenceScore.toLowerCase());
    }

    // patchValue() above already fires developmentStatus's valueChanges
    // (which re-runs updateEstimatedFieldsValidators()), but call it
    // explicitly too so the Estimated Volume/Dollars requiredness is
    // guaranteed correct for this concept's actual status regardless of
    // subscription timing/ordering.
    this.updateEstimatedFieldsValidators(c.DevelopmentStatus ?? 'New');

    // The backend's getConcept response includes IdeationRequestorName /
    // DataScienceProgrammerName directly on the concept, but the dropdowns
    // only resolve a display label by matching IdeationRequestorId /
    // DataScienceProgrammerId against ideationRequestorOptions /
    // dataScienceProgrammerOptions — separately-loaded master-data lists.
    // If the assigned person isn't in that list (deactivated, filtered out,
    // or master data simply hasn't loaded yet), the <select> renders blank
    // even though the API clearly returned their name. Stash the concept's
    // own id/name here so ensureAssignedUsersVisible() can (re)inject them
    // into the options lists — both now, and again once loadMasterData()
    // resolves, since that call wholesale-replaces these arrays and would
    // otherwise wipe the fallback entry back out (same race documented on
    // prefillIdeationRequestor()).
    this.lastConceptRequestor = c.IdeationRequestorId
      ? { id: Number(c.IdeationRequestorId), name: c.IdeationRequestorName }
      : null;
    this.lastConceptProgrammer = c.DataScienceProgrammerId
      ? { id: Number(c.DataScienceProgrammerId), name: c.DataScienceProgrammerName }
      : null;
    this.ensureAssignedUsersVisible();
  }

  /** Holds the currently-loaded concept's assigned Ideation Requestor /
   *  Data Science Programmer (id + name straight from getConcept), so they
   *  can be re-injected into the dropdown options list whenever it's
   *  (re)loaded — see ensureAssignedUsersVisible(). null when nobody's
   *  assigned or no concept is loaded (create flow). */
  private lastConceptRequestor: { id: number; name: string } | null = null;
  private lastConceptProgrammer: { id: number; name: string } | null = null;

  /** Makes sure the loaded concept's assigned Ideation Requestor / Data
   *  Science Programmer always appear in their dropdowns with the correct
   *  name, even if they're missing from the master-data options list
   *  (deactivated, filtered out, or master data hasn't loaded yet). Called
   *  from patchForm() and again from loadMasterData()'s callback, since
   *  master data can resolve either before or after the concept does. */
  private ensureAssignedUsersVisible(): void {
    if (this.lastConceptRequestor) {
      this.ensureOptionPresent(
        this.ideationRequestorOptions,
        this.lastConceptRequestor.id,
        this.lastConceptRequestor.name,
        'Ideation Requestor'
      );
    }
    if (this.lastConceptProgrammer) {
      this.ensureOptionPresent(
        this.dataScienceProgrammerOptions,
        this.lastConceptProgrammer.id,
        this.lastConceptProgrammer.name,
        'Data Science Programmer'
      );
    }
  }

  /** Adds a { id, name, role_name } entry to a dropdown's options array if
   *  that id isn't already present — used so a concept's saved Ideation
   *  Requestor / Data Science Programmer always shows their name, even if
   *  they're missing from the master-data list the dropdown was populated
   *  from. No-op when id is falsy/0 (nobody assigned) or already present. */
  private ensureOptionPresent(
    options: { id: number; name: string; role_name?: string }[],
    id: number | null | undefined,
    name: string | null | undefined,
    fallbackRole: string
  ): void {
    if (!id) return;
    const exists = options.some(u => Number(u.id) === Number(id));
    if (!exists) {
      options.push({ id: Number(id), name: name || `User #${id}`, role_name: fallbackRole });
    }
  }

  private patchMeta(c: any): void {
    if (c.CreatedDate) this.createdDate = new Date(c.CreatedDate);
    if (c.UpdatedDate) this.updatedDate = new Date(c.UpdatedDate);

    if (c.ConceptId) {
      this.conceptId = c.ConceptId;
    }
    this.displayConceptId = c.CurrentConceptId ?? this.conceptId;

    this.developmentCompleted = c.DevelopmentCompleted ? 1 : 0;
    this.clientApprovalCompleted = c.ClientApprovalCompleted ? 1 : 0;
    this.supportingDocumentsCompleted = c.SupportingDocumentsCompleted ? 1 : 0;
    this.isDraftConcept = !!c.isDraft || c.RecordType === 'DRAFT';

  }

  private patchAttachments(files: any[]): void {
    this.attachments = { specs: [], table: [], other: [], approval: [] };

    files
      // Supporting Documents share this same active_files array (tagged
      // AttachmentType: "supporting_docs") but belong in their own
      // section — see patchSupportingDocs() below, not here.
      .filter(f => (f.AttachmentType ?? '').toLowerCase() !== 'supporting_docs')
      .forEach(f => {
        const rawCat = (f.AttachmentType ?? '').toLowerCase() as AttachCategory;
        const cat: AttachCategory = ['specs', 'table', 'other', 'approval'].includes(rawCat) ? rawCat : 'other';

        this.attachments[cat] = [
          ...this.attachments[cat],
          {
            id: f.AttachmentId ? `att_${f.AttachmentId}` : this.generateAttachId(),
            name: f.FileName,
            size: f.FileSize ?? 0,
            progress: 100,
            file: new File([], f.FileName),
            // GET /api/download-attachment/{attachment_id} — streams the file
            // back with Content-Disposition: inline, so the blob just gets
            // handed to the in-app Word/Excel/PDF viewer (see viewAttachment).
            downloadUrl: f.AttachmentId
              ? `api/download-attachment/${f.AttachmentId}`
              : undefined,
            attachmentId: f.AttachmentId ?? undefined
          }
        ];
      });
  }

  /** Restores the Supporting Documents cards from the same active_files
   *  array patchAttachments() reads — the two sections share one backend
   *  table, distinguished only by AttachmentType. supportingDocs is a
   *  fixed 3-slot layout, so saved docs fill slots in order and any
   *  remaining slots stay blank/ready for upload. */
  private patchSupportingDocs(files: any[]): void {
    const savedDocs = files
      .filter(f => (f.AttachmentType ?? '').toLowerCase() === 'supporting_docs')
      // Sort by DocIndex — the backend's explicit slot-position field for
      // this doc (0, 1, 2…), set once when the doc is first added to a
      // slot and never reshuffled by later edits to that same doc. This is
      // the correct ordering key; AttachmentId (creation order) was used
      // before and mostly lines up with DocIndex, but isn't guaranteed to
      // (e.g. after a doc is removed and a new one takes over the freed
      // slot, the new one gets a fresh/high AttachmentId but keeps the old
      // slot's DocIndex). Falls back to AttachmentId only if DocIndex is
      // ever missing from a record. Without a stable sort here, the
      // backend can return this list re-ordered by last-updated timestamp
      // — e.g. editing Doc 1's sourceurl makes Doc 1 "most recently
      // updated" and bumps its position in the response — which is what
      // made Doc 1 and Doc 2's content appear to swap cards on every
      // resubmit.
      .sort((a, b) => (a.DocIndex ?? a.AttachmentId ?? 0) - (b.DocIndex ?? b.AttachmentId ?? 0))
      .map(f => {
        // A doc is URL-only when sourceurl is filled but FileSize is 0 (or
        // the backend stored no real bytes — e.g. the user only pasted a link).
        const hasRealFile = f.FileSize > 0;

        return {
          name: f.DocName ?? f.FileName ?? '',
          sourceurl: f.sourceurl ?? '',
          pdfLocation: '',
          uploadProgress: hasRealFile ? 100 : 0,
          // Keep file null for URL-only docs so the UI correctly shows
          // the link state instead of a false "Uploaded" success banner.
          file: hasRealFile
            ? new File([], f.FileName)   // placeholder for viewer
            : null,
          downloadUrl: hasRealFile && f.AttachmentId
            ? `api/download-attachment/${f.AttachmentId}`
            : undefined,
          restored: true,
          // Reloaded/restored docs (including the soft-reload right after a
          // successful submit) must NOT show the success banner — it's only
          // meant to flash for 5s at the moment of a fresh upload, not every
          // time this list gets rebuilt from the backend afterward.
          successBannerVisible: false,
          // Remembered so an unmodified resubmit can still tell the backend
          // this slot's file is intact — see resolveDocFileMeta() in
          // onDocSubmit(). Without this, a slot that already has a real
          // file looks file-less the moment it's saved again without the
          // user picking a new file.
          originalFileName: hasRealFile ? f.FileName : undefined,
          originalFileSize: hasRealFile ? f.FileSize : undefined,
          // Needed by removeSupportingDoc() to call the delete API — a
          // restored slot always has a real backend record, even URL-only
          // ones (they're still a row in active_files).
          attachmentId: f.AttachmentId ?? undefined,
          // STABLE slot identity — read straight off the backend record,
          // NEVER derived from this doc's position in `savedDocs`/the sort
          // above (which is only for display ordering and can legitimately
          // reorder relative to raw array position). Falls back to
          // AttachmentId only for legacy rows that predate this column.
          docIndex: f.DocIndex ?? f.AttachmentId ?? 0
        } as SupportingDoc;
      });

    this.supportingDocs = savedDocs.length > 0
      ? [...savedDocs]
      : [this.blankDoc(), this.blankDoc(), this.blankDoc()];

    // Make sure the next brand-new slot's docIndex can never collide with
    // any DocIndex just restored from the backend (including ones that
    // belong to docs the user hasn't loaded here, if any exist higher).
    const maxRestoredIndex = this.supportingDocs.reduce(
      (max, d) => Math.max(max, d.docIndex ?? -1), -1
    );
    this.nextDocIndex = Math.max(this.nextDocIndex, maxRestoredIndex + 1);

    // Snapshot what was just restored so onDocSubmit()'s "did anything
    // change" guard has a baseline to compare against — this runs on every
    // (re)load of this tab, including the soft-refresh right after a
    // successful Supporting Documents submit, so the snapshot always
    // reflects whatever is currently persisted.
    this.snapshotSupportingDocs();
  }

  /** Serializable signature of the currently-saved-or-saveable Supporting
   *  Documents state (same shape/filter onDocSubmit() sends as validDocs),
   *  used by getSupportingDocsChanged() to detect a no-op resubmit. Keyed by
   *  docIndex (stable slot identity, never array position — see docIndex's
   *  doc comment) and sorted by it, so unrelated reordering in the array
   *  never looks like a change. */
  private originalSupportingDocsSnapshot = '';

  private buildSupportingDocsSignature(): string {
    return JSON.stringify(
      this.supportingDocs
        .filter(d => d.file || d.sourceurl?.trim() || d.pdfLocation?.trim())
        .map(d => {
          const { fileName, fileSize } = this.resolveDocFileMeta(d);
          return {
            docIndex: d.docIndex,
            name: d.name || '',
            sourceurl: d.sourceurl || '',
            pdfLocation: d.pdfLocation || '',
            fileName: fileName || '',
            fileSize: fileSize || 0,
          };
        })
        .sort((a, b) => a.docIndex - b.docIndex)
    );
  }

  private snapshotSupportingDocs(): void {
    this.originalSupportingDocsSnapshot = this.buildSupportingDocsSignature();
  }

  /** True if the Supporting Documents tab's content differs from what was
   *  last saved/loaded — mirrors getChangedWatchedFields()/getAttachmentsChanged()
   *  but for this tab's own array-backed (not form-backed) state. Used by
   *  onDocSubmit()'s "did anything change" guard. */
  private getSupportingDocsChanged(): boolean {
    return this.buildSupportingDocsSignature() !== this.originalSupportingDocsSnapshot;
  }

  // ── Development Notes (from API) ──────────────────────────────────────
  /** TODO: confirm these field names against your actual development_notes
   *  record shape — the fallbacks below just guard against a couple of
   *  likely naming conventions (e.g. "Createdby" elsewhere in this payload
   *  is a numeric user id, not a name, so we fall back to "User #<id>"
   *  if no display name is supplied). */
  private patchDevNotes(notes: any[]): void {
    const avatarPalette = [
      '#6366f1',
      '#8b5cf6',
      '#0ea5e9',
      '#10b981',
      '#f59e0b',
      '#ef4444'
    ];

    const userColorMap = new Map<string, string>();

    const getAvatarColor = (author: string): string => {
      if (!userColorMap.has(author)) {
        const colorIndex = userColorMap.size % avatarPalette.length;
        userColorMap.set(author, avatarPalette[colorIndex]);
      }
      return userColorMap.get(author)!;
    };

    this.devNotes = (notes ?? []).map((n: any, i: number) => {
      const author =
        n.AuthorName ??
        n.CreatedByName ??
        n.author ??
        (n.Createdby ? `User #${n.Createdby}` : 'User');

      const text = n.NoteText ?? n.Note ?? n.text ?? '';
      const date = n.CreatedDate ?? n.createdDate ?? n.time;

      return {
        id: n.NoteId ? `note_${n.NoteId}` : `note_${i}_${Date.now()}`,
        author,
        initials: author.charAt(0).toUpperCase(),
        avatarBg: getAvatarColor(author),
        time: date ?? '',
        text,
        RoleName: n.RoleName ?? n.Role ?? '',
        persisted: true
      };
    });

    this.scrollNotesToBottom();
  }

  private lastSyncedClientEstimated: { volume: any; dollars: any } = { volume: '', dollars: '' };


  // ── Client Approval (from API) ────────────────────────────────────────
  /** TODO: confirm these field names against your actual client_approvals
   *  record shape — mapped here to mirror the PascalCase convention used
   *  elsewhere in this payload (ConceptName, DevelopmentStatus, etc.).
   *  If a concept has no approval submissions yet, client_approvals will
   *  be empty and these fields are reset to blank.
   *
   *  IMPORTANT: this must always patchValue, even when approvals is
   *  empty — the component instance is reused across concepts (see the
   *  routeSub comment in ngOnInit), so an early return here used to
   *  leave the PREVIOUS concept's Client Approval fields sitting in the
   *  form untouched. Switching to a concept with no approval yet then
   *  showed that leftover data as if it belonged to the new concept. */
  private patchClientApproval(
    approvals: any[],
    isConceptSwitch: boolean = true,
    developmentEstimated: { volume: any; dollars: any } = { volume: '', dollars: '' }
  ): void {
    const latest = approvals && approvals.length > 0
      ? [...approvals].sort((a, b) =>
        new Date(b.CreatedDate ?? b.SubmittedDate ?? 0).getTime() -
        new Date(a.CreatedDate ?? a.SubmittedDate ?? 0).getTime()
      )[0]
      : null;

    const clientApprovalValues: Record<string, any> = {
      clientConceptName: latest?.ClientConceptName ?? '',
      clientConceptDescription: latest?.ClientConceptDescription ?? '',
      clientApprovalStatus: latest?.ClientApprovalStatus ?? '',
      submittedToClientOn: latest?.SubmittedToClientOn
        ? latest.SubmittedToClientOn.split('T')[0]
        : '',
      clientApprovalNotes: latest?.ClientApprovalNotes ?? '',
      // Prefer a value actually recorded against a client approval
      // submission (in case the backend keeps its own per-approval
      // figure); fall back to the Development tab's currently-saved
      // value for a concept that has no approval submission yet.
      clientEstimatedVolume: latest?.EstimatedVolume ?? developmentEstimated.volume ?? '',
      clientEstimatedDollars: latest?.EstimatedDollars ?? developmentEstimated.dollars ?? ''
    };

    if (isConceptSwitch) {
      // Genuine navigation to a different concept (or this concept's
      // very first load). This component instance is reused across
      // concepts — see routeSub in ngOnInit — so the Client Approval
      // tab must be fully reset to exactly what the server has for THIS
      // concept, even if that means blanking fields: any leftover value
      // here would otherwise be mistaken for this concept's own data.
      this.form.patchValue(clientApprovalValues);
    } else {
      // Same-concept soft refresh (refreshConceptData(), fired right
      // after some OTHER tab's submit succeeds — e.g. Supporting
      // Documents). Client Approval may not have been submitted yet, so
      // res.client_approvals can be stale/empty relative to whatever the
      // user has currently typed into these controls. Only re-sync
      // controls the user hasn't touched since the last load/submit —
      // overwriting a dirty control here would silently erase
      // unsubmitted work and, since the control is already
      // touched/dirty, instantly show a red "required" error on a field
      // the user just filled in.
      for (const name of Object.keys(clientApprovalValues)) {
        if (name === 'clientEstimatedVolume' || name === 'clientEstimatedDollars') {
          // Handled separately below — .dirty alone isn't a reliable
          // signal for these two (see lastSyncedClientEstimated).
          continue;
        }
        const control = this.form.get(name);
        if (control && !control.dirty) {
          control.setValue(clientApprovalValues[name]);
        }
      }
      const volControl = this.form.get('clientEstimatedVolume');
      const dolControl = this.form.get('clientEstimatedDollars');
      const volUnchangedSinceSync =
        String(volControl?.value ?? '').trim() === String(this.lastSyncedClientEstimated.volume ?? '').trim();
      const dolUnchangedSinceSync =
        String(dolControl?.value ?? '').trim() === String(this.lastSyncedClientEstimated.dollars ?? '').trim();

      if (volUnchangedSinceSync) {
        volControl?.setValue(clientApprovalValues['clientEstimatedVolume']);
      }
      if (dolUnchangedSinceSync) {
        dolControl?.setValue(clientApprovalValues['clientEstimatedDollars']);
      }
    }
    this.lastSyncedClientEstimated = {
      volume: this.form.get('clientEstimatedVolume')?.value ?? '',
      dollars: this.form.get('clientEstimatedDollars')?.value ?? ''
    };

    // The dirty-flag re-derivation below assumes we just did a full,
    // authoritative reset of the Client Approval group (the
    // isConceptSwitch branch above) — on a soft refresh we've already
    // left any in-progress dirty state exactly as the user made it, so
    // there's nothing to re-derive.
    if (!isConceptSwitch) return;

    // Re-derive whether clientConceptName counts as "customized" for
    // THIS freshly-loaded concept, rather than carrying over whatever
    // dirty flag a previously-viewed concept left behind (this component
    // instance is reused across concepts — see the routeSub comment in
    // ngOnInit). A saved value that actually differs from the concept's
    // own name is a genuine customization made in an earlier session, so
    // mark it dirty to protect it from being auto-overwritten by
    // ensureClientConceptName(); otherwise mark pristine so it stays
    // eligible to auto-sync from conceptName going forward.
    const clientNameControl = this.form.get('clientConceptName');
    const conceptNamePersisted = this.savedConceptName.trim();
    const savedClientName = (latest?.ClientConceptName ?? '').trim();
    if (savedClientName && savedClientName !== conceptNamePersisted) {
      clientNameControl?.markAsDirty();
    } else {
      clientNameControl?.markAsPristine();
    }
  }

  // ── Master data (dropdown sources) ────────────────────────────────────
  private loadMasterData(): void {
    this.masterDataLoading = true;

    this.service.getmasterdata().subscribe({
      next: (res) => {
        const data = res?.data ?? {};
        this.developmentStatusOptions = data.development_status ?? [];
        this.priorityOptions = data.priority_status ?? [];
        this.clientOptions = data.clients ?? [];
        this.masterConceptOptions = data.master_concepts ?? [];
        this.reviewTypeOptions = data.review_types ?? [];
        this.claimTypeOptions = data.claim_types ?? [];
        this.ideationRequestorOptions = data.ideation_requestors ?? [];
        this.dataScienceProgrammerOptions = data.datascience_programmers ?? [];
        this.clientApprovalstatusOptions = data.ClientApproval_status ?? [];
        this.masterDataLoading = false;

        // Master data (and therefore ideationRequestorOptions) loads
        // asynchronously and can resolve after resetToNewConcept() already
        // ran on first page load — try the auto-fill again now that the
        // options list actually has entries to match against.
        this.prefillIdeationRequestor();

        // Same race, but for an already-loaded concept's assigned users:
        // loadMasterData() just replaced ideationRequestorOptions /
        // dataScienceProgrammerOptions wholesale, which would silently
        // drop the fallback entry patchForm() injected earlier if this
        // resolves afterward. Re-inject it now.
        this.ensureAssignedUsersVisible();

        // Same race: refreshAllowedStatuses() may have run before
        // developmentStatusOptions was populated, in which case its
        // filter against an empty array produced an empty dropdown.
        // Re-run it now against whatever status is currently on the form.
        const currentStatus = this.lastKnownDevStatus || 'New';

        if (this.isReadOnly || !Number(sessionStorage.getItem('userId'))) {
          // Read-only roles never populate this.allowedStatuses (they skip
          // the getAllowedStatuses call entirely) — rebuild directly from
          // developmentStatusOptions, which just became available.
          this.allowedStatusOptions = this.developmentStatusOptions.filter(
            d => d.value === currentStatus
          );
        } else if (this.allowedStatuses.length) {
          this.applyAllowedStatusFilter(currentStatus);
        }
      },
      error: (err) => {
        console.error('Failed to load master data:', err);
        this.toastr.error('Failed to load dropdown options', 'Error');
        this.masterDataLoading = false;
      }
    });
  }

  /** status name -> requiresEstimates, sourced ENTIRELY from
   *  GET /api/allowed-statuses (see refreshAllowedStatuses() below).
   *
   *  This used to be a hardcoded EARLY_STAGE_STATUSES string array that
   *  had to be kept in sync BY HAND with status_permissions.py's
   *  run_stage_validations() (which gates on numeric status id >= 4).
   *  Those were two independent sources of truth for the exact same
   *  rule, and they could silently drift — e.g. a new status added to
   *  the backend that nobody remembered to also add to this array.
   *  Now the backend computes requires_estimates() once and the
   *  frontend just reads it off the response; there is nothing here
   *  left to keep in sync when a status is added or renumbered.
   *
   *  '' (no status chosen yet, brand-new concept) defaults to false —
   *  same as today's "New" behavior — until the first
   *  /api/allowed-statuses response arrives. */
  private statusRequiresEstimatesMap: Record<string, boolean> = { '': false };

  private setStatusRequiresEstimates(status: { name?: string; value?: string } & { requiresEstimates: boolean }): void {
    const name = status?.name ?? (status as any)?.value;
    if (name) {
      this.statusRequiresEstimatesMap[name] = !!status.requiresEstimates;
    }
  }

  /** Drives the "*" required-mark next to Estimated Volume / Estimated
   *  Dollars on the Concept Development tab — true once Development
   *  Status has advanced to "Result Set QA" or beyond (per the backend's
   *  requires_estimates()), matching updateEstimatedFieldsValidators()
   *  below so the mark never shows when the field is actually optional
   *  (and vice versa). */
  get estimatedFieldsRequired(): boolean {
    const status = this.form.get('developmentStatus')?.value ?? '';
    return this.statusRequiresEstimatesMap[status] ?? false;
  }

  /** Applies/removes the "required" validator on Estimated Volume /
   *  Estimated Dollars based on where the concept currently sits in
   *  Development Status, per statusRequiresEstimatesMap above.
   *  Validators.min(1) stays on either way, so a value that IS entered
   *  still can't be zero or negative. Called on form build, whenever
   *  Development Status changes, and again once refreshAllowedStatuses()
   *  populates the map (see the race note there), so the rule always
   *  reflects the concept's current status rather than whatever status
   *  was on the form last, or a stale/default map entry. */
  private updateEstimatedFieldsValidators(status: string | null | undefined): void {
    const requiresEstimates = this.statusRequiresEstimatesMap[status ?? ''] ?? false;
    const validators = requiresEstimates
      ? [Validators.required, Validators.min(1)]
      : [Validators.min(1)];

    ['estimatedVolume', 'estimatedDollars'].forEach(name => {
      const control = this.form.get(name);
      control?.setValidators(validators);
      control?.updateValueAndValidity({ emitEvent: false });
    });
  }

  // ── Form ──────────────────────────────────────────────────────────────
  private buildForm(): void {
    // A brand-new/blank form is never Production-locked — reset here so
    // resetToNewConcept() (which calls buildForm()) doesn't carry over a
    // previously-loaded concept's locked state onto a fresh one.
    this.currentConceptStatus = '';
    this.dismissProductionLockBanner();

    this.form = this.fb.group({
      conceptName: ['', Validators.required],
      clientName: ['', Validators.required],
      masterConceptName: ['', Validators.required],
      reviewType: ['', Validators.required],
      claimType: ['', Validators.required],
      developmentStatus: [''],
      priority: [''],
      haloNumber: [''],
      Internalconceptdescription: [''],
      estimatedVolume: [null, [Validators.min(1)]],
      estimatedDollars: ['', [Validators.min(1)]],
      confidenceScore: this.fb.group({ value: ['low'] }),
      ideationRequestor: [''],
      dataScienceProgrammer: [''],
      previousReportId: [''],
      qaSchedule: ['', [ConceptCreateComponent.validDateRange, ConceptCreateComponent.notPastDate]],
      productionSchedule: ['', [ConceptCreateComponent.validDateRange, ConceptCreateComponent.notPastDate]],


      // Client Approval
      clientConceptName: ['', Validators.required],
      clientConceptDescription: ['', Validators.required],
      clientApprovalStatus: ['', Validators.required],
      submittedToClientOn: ['', [Validators.required, ConceptCreateComponent.validDateRange, ConceptCreateComponent.notPastDate]],
      clientApprovalNotes: ['', Validators.required],
      clientEstimatedVolume: [null, [Validators.required, Validators.min(1)]],
      clientEstimatedDollars: ['', [Validators.required, Validators.min(1)]]
    });

    // Start with the correct (optional, since a fresh form has no status
    // chosen yet -> falls into EARLY_STAGE_STATUSES' '' entry) requiredness,
    // then keep it in sync any time the user changes Development Status
    // themselves via the dropdown.
    this.updateEstimatedFieldsValidators(this.form.get('developmentStatus')?.value);
    this.form.get('developmentStatus')?.valueChanges.subscribe(status => {
      this.updateEstimatedFieldsValidators(status);
    });
  }

  // ── Tabs ──────────────────────────────────────────────────────────────
  setActiveTab(key: string): void {
    this.activeTab = key;

    // Whenever the user lands on Client Approval, make sure the
    // client-facing "Concept Name" field has something in it. If it's
    // still blank (brand-new concept, or an existing one that never had
    // a client name set), fall back to the Concept Development name so
    // the field never appears empty.
    if (key === 'development') {
      // this.loadUserFiles();
    }
    if (key === 'approval') {
      this.ensureClientConceptName();
    }
    if (key === 'documents') {
      // this.loadUserFiles();
    }
  }

  /** Keeps clientConceptName mirroring the last-PERSISTED conceptName
   *  (savedConceptName — see its declaration) for as long as the user
   *  hasn't actually customized clientConceptName themselves. Two things
   *  this deliberately gets right:
   *
   *  1. Only checking "is it blank" (the old behavior) breaks the moment
   *     it's auto-filled once: it's no longer blank, so a later edit to
   *     the main Concept Name (e.g. "concept-28" -> "concept-28.1") would
   *     never propagate here, even though the user never typed anything
   *     of their own into this field. `dirty` is the right signal
   *     instead — it's true only once the user has actually edited
   *     clientConceptName directly (this field uses formControlName, so
   *     Angular sets `dirty` on real user input; programmatic setValue()
   *     calls, like this one, never do). Once dirty,
   *     it's a deliberately customized client-facing name and is left
   *     alone.
   *  2. It reads savedConceptName, not the live conceptName control
   *     value — an in-progress edit to Concept Name on the Development
   *     tab that hasn't been saved yet must not leak over to the Client
   *     Approval tab. Only once that edit is actually saved does
   *     savedConceptName (and therefore this sync) pick it up — see
   *     patchForm(), which sets savedConceptName from whatever the
   *     backend just confirmed as persisted.
   *
   *  Called when the Approval tab opens and again whenever the user
   *  clears the field on blur. */
  private isClientConceptNameCustomized(): boolean {
    const current = (this.form.get('clientConceptName')?.value ?? '').toString().trim();
    return !!current && current !== this.savedConceptName.trim();
  }

  private ensureClientConceptName(): void {
    const control = this.form.get('clientConceptName');
    if (!this.isClientConceptNameCustomized()) {
      control?.setValue(this.savedConceptName, { emitEvent: false });
    }
  }

  /** If the client clears the Concept Name field entirely, revert it to
   *  the last-persisted concept name (savedConceptName, not the live
   *  conceptName control — same reasoning as ensureClientConceptName())
   *  on blur instead of leaving it blank, while still letting them
   *  freely edit it in between. Unlike ensureClientConceptName(), this
   *  always reverts when blank — regardless of dirty — since an emptied
   *  field has no customization left to preserve. Reverting here also
   *  resets dirty back to false: the field is back to purely mirroring
   *  savedConceptName, so it should resume auto-syncing on future saves
   *  until the user actually customizes it again. */
  onClientConceptNameBlur(): void {
    const control = this.form.get('clientConceptName');
    const value = control?.value?.trim();
    if (!value) {
      control?.setValue(this.savedConceptName, { emitEvent: false });
      control?.markAsPristine();
    }
  }

  onTabChange(key: string): void {
    // Draft concepts only have the Concept Development tab — clicking
    // Client Approval or Supporting Document shouldn't switch tabs, it
    // should explain why, via a dismissible banner instead of leaving the
    // user to guess from a disabled-looking button.
    if ((!this.conceptId || this.isDraftConcept) && key !== 'development') {
      this.flashDraftLockBanner();
      return;
    }

    this.showDraftLockBanner = false;
    this.activeTab = key;
    if (key === 'approval') {
      // Pre-fill clientConceptName from conceptName if blank,
      // so the field is never empty when the tab loads.
      this.ensureClientConceptName();
      this.cdr.detectChanges();
    }
  }

  /** Shows the draft-lock banner and auto-dismisses it after a few
   *  seconds. Restarts the timer on repeated clicks so it doesn't
   *  disappear mid-read if the user clicks more than once. */
  private flashDraftLockBanner(): void {
    this.showDraftLockBanner = true;
    if (this.draftLockBannerTimer) {
      clearTimeout(this.draftLockBannerTimer);
    }
    this.draftLockBannerTimer = setTimeout(() => {
      this.showDraftLockBanner = false;
      this.draftLockBannerTimer = null;
    }, 6000);
  }

  dismissDraftLockBanner(): void {
    this.showDraftLockBanner = false;
    if (this.draftLockBannerTimer) {
      clearTimeout(this.draftLockBannerTimer);
      this.draftLockBannerTimer = null;
    }
  }

  /** Shows the Production-lock banner and auto-dismisses it after a few
   *  seconds — same timed flash-in/out as flashDraftLockBanner() above,
   *  triggered once when a Production-locked concept finishes loading
   *  (see patchForm()) rather than staying pinned on screen the whole
   *  time the tab is open. */
  private flashProductionLockBanner(): void {
    this.showProductionLockBanner = true;
    if (this.productionLockBannerTimer) {
      clearTimeout(this.productionLockBannerTimer);
    }
    this.productionLockBannerTimer = setTimeout(() => {
      this.showProductionLockBanner = false;
      this.productionLockBannerTimer = null;
    }, 6000);
  }

  dismissProductionLockBanner(): void {
    this.showProductionLockBanner = false;
    if (this.productionLockBannerTimer) {
      clearTimeout(this.productionLockBannerTimer);
      this.productionLockBannerTimer = null;
    }
  }

  // ── Confidence Score ──────────────────────────────────────────────────
  /** Confidence Score is outside the Data Science Programmer's allowed
   *  field list (and is fully locked for read-only roles). The toggle
   *  buttons in the template call this directly on click, so form.disable()
   *  alone won't stop it — this guard is the actual enforcement point.
   *  Backed up by [disabled] on the buttons themselves (see template). */
  get canEditConfidenceScore(): boolean {
    if (this.isRecordLocked) return false;
    return this.canFullEdit;
  }

  setConfidenceScore(level: 'low' | 'medium' | 'high'): void {
    if (!this.canEditConfidenceScore) return;
    this.form.get('confidenceScore.value')?.setValue(level);
  }

  /** Drafts intentionally skip most of the frontend's required-field
   *  checks (see onSaveAsDraft()), but several columns are still NOT NULL
   *  on the backend. When one of those is missing, the API can bubble up
   *  a raw SQL/DB error (e.g. "NOT NULL constraint failed: ..." or similar
   *  driver text) instead of a clean validation message. Showing that raw
   *  text to the user is confusing and looks like the app is broken — this
   *  is almost always actually just a missing required field, so map any
   *  DB/SQL-shaped error message to the same friendly message a frontend
   *  validation failure would show, and only fall back to the raw text for
   *  errors that clearly aren't about missing data. */
  private friendlyErrorMessage(raw: string): string {
    const sqlErrorPattern = /sql|syntax error|constraint|column|not null|database|db error|exception|cannot insert|violates|duplicate key|foreign key/i;
    if (raw && sqlErrorPattern.test(raw)) {
      return 'Please fill in all required fields and try again.';
    }
    return raw;
  }

  /** Shared "is this required field effectively empty" check used by both
   *  onSubmit() and onApprovalSubmit(). estimatedVolume/estimatedDollars
   *  get extra treatment: Validators.required alone doesn't reject 0 (it's
   *  neither null, undefined, nor an empty string), but 0 isn't a
   *  meaningful estimate — a concept can't be submitted claiming 0 volume
   *  or 0 dollars, so those two controls treat any value <= 0 as missing
   *  too, on top of the ordinary blank check every other field gets. */
  private isRequiredFieldMissing(controlName: string, value: any): boolean {
    if (value === null || value === undefined || String(value).trim() === '') {
      return true;
    }
    if (
      (controlName === 'estimatedVolume' || controlName === 'estimatedDollars' ||
        controlName === 'clientEstimatedVolume' || controlName === 'clientEstimatedDollars') &&
      Number(value) <= 0
    ) {
      return true;
    }
    return false;
  }

  /** Numeric-typed columns on the backend (estimatedVolume, estimatedDollars,
   *  haloNumber, previousReportId) reject an empty string outright — SQL
   *  Server can't implicitly convert '' to numeric and throws a raw ODBC
   *  error ("Error converting data type nvarchar to numeric") instead of a
   *  clean validation message. This mostly bites Save as Draft, where these
   *  fields are allowed to be left blank: a blank numeric control's value
   *  is '' (see buildForm()), not null, so it has to be coerced here before
   *  going into the metadata payload — sending null (which becomes SQL
   *  NULL) instead of '' is what the column actually accepts. */
  private numericOrNull(value: any): number | null {
    if (value === null || value === undefined || String(value).trim() === '') {
      return null;
    }
    const n = Number(value);
    return isNaN(n) ? null : n;
  }

  // ── Save / Submit ─────────────────────────────────────────────────────
  onSave(): void {
    if (this.form.valid) {
      console.log('Saved:', this.form.value);
      this.updatedDate = new Date();
    } else {
      this.form.markAllAsTouched();
    }
  }
  public get recordLockedMessage(): string {
    if (this.isSupersededLocked) {
      return 'This concept is Superseded. No changes can be made.';
    }
    return 'This concept is in Production and is locked for editing.';
  }

  async onSubmit(): Promise<void> {
    // Flips the Concept Development tab's own "submit was attempted"
    // flag — the template uses this (not shared-control .touched) to
    // decide whether to show estimatedVolume/estimatedDollars errors, so
    // a failed submit here can never leak an error onto Client Approval.
    this.developmentSubmitAttempted = true;
    if (!this.canEdit) {
      this.toastr.error('You do not have permission to submit concepts.', 'Access Denied');
      return;
    }
    // Safety net behind canSaveOrCreate/the [disabled] binding on the
    // button — Data Science Programmer never originates a new concept
    // (only "Update"s existing ones assigned to them), so the "Create
    // Concept" action stays blocked for that role even if triggered some
    // other way while the concept doesn't exist yet.
    if (this.canDSEdit && !this.developmentCompleted) {
      this.toastr.error('You do not have permission to create a new concept.', 'Access Denied');
      return;
    }
    // Safety net behind canSubmitForm/the [disabled] binding on the
    // button — once Production-locked, only Manager may still submit
    // (and only to persist a Development Status change).
    if (this.isRecordLocked && !(this.isProductionLocked && this.isManager)) {
      this.toastr.error(this.recordLockedMessage, 'Access Denied');
      return;
    }
    // Safety net behind the [disabled] binding on the button itself —
    // blocks the call even if it's triggered some other way (e.g. Enter
    // key) while a file's progress bar hasn't reached 100% yet.
    if (this.isAttachmentUploading) {
      this.toastr.error('Please wait for all files to finish uploading.', 'Upload in progress');
      return;
    }

    const qaControl = this.form.get('qaSchedule');
    const prodControl = this.form.get('productionSchedule');

    if (qaControl?.hasError('invalidDateRange')) {
      this.toastr.error('Please enter a valid QA Schedule date.', 'Invalid Date');
      qaControl.markAsTouched();
      return;
    }

    if (prodControl?.hasError('invalidDateRange')) {
      this.toastr.error('Please enter a valid Production Schedule date.', 'Invalid Date');
      prodControl.markAsTouched();
      return;
    }
    // Client Name / Master Concept Name / Review Type / Claim Type are
    // only required at creation time — once the concept exists they're
    // hidden (see template) and updates are identified by concept_id
    // alone, so they're skipped from required-field validation here too.
    // Everything else that's required (Concept Name, Estimated Volume,
    // Estimated Dollars) stays editable and mandatory for the entire
    // life of the concept, so clearing one of those on an update must
    // still block submission instead of silently saving it blank.
    const requiredFields: { control: string; label: string }[] = [
      { control: 'conceptName', label: 'Concept Name' },
      // Estimated Volume / Estimated Dollars are only mandatory from
      // "Result Set QA" onward (see estimatedFieldsRequired /
      // updateEstimatedFieldsValidators) - checking them here
      // unconditionally used to block submission of a brand-new
      // (early-stage) concept even though the reactive validators on
      // the controls themselves correctly treated them as optional.
      ...(this.estimatedFieldsRequired
        ? [
          { control: 'estimatedVolume', label: 'Estimated Volume' },
          { control: 'estimatedDollars', label: 'Estimated Dollars' },
        ]
        : []),
      ...(!this.conceptId
        ? [
          { control: 'clientName', label: 'Client Name' },
          { control: 'masterConceptName', label: 'Master Concept Name' },
          { control: 'reviewType', label: 'Review Type' },
          { control: 'claimType', label: 'Claim Type' },
        ]
        : []),
    ];

    const missing = requiredFields.filter(f => {
      const value = this.form.get(f.control)?.value;
      return this.isRequiredFieldMissing(f.control, value);
    });

    // Estimated Volume / Estimated Dollars carry Validators.min(1) on the
    // control itself regardless of estimatedFieldsRequired (see
    // buildForm()/updateEstimatedFieldsValidators()), so a value the user
    // DID type in must still be > 0 even at an early-stage status where
    // the field is otherwise optional and therefore isn't in
    // requiredFields above. onSubmit() never checks this.form.valid, so
    // without this explicit check that min(1) violation is silently
    // ignored and a 0 gets sent straight to the backend.
    const invalidEstimates = ['estimatedVolume', 'estimatedDollars'].filter(
      name => this.form.get(name)?.hasError('min')
    );

    const dateFields = ['qaSchedule', 'productionSchedule', 'submittedToClientOn'];
    const badDate = dateFields.find(f =>
      this.form.get(f)?.hasError('invalidDateRange') || this.form.get(f)?.hasError('pastDate')
    );
    if (badDate) {
      const isPast = this.form.get(badDate)?.hasError('pastDate');
      this.toastr.error(
        isPast
          ? 'QA Schedule and Production Schedule cannot be set in the past.'
          : 'Please enter a valid date (year must be between 1900 and 2100).',
        'Invalid Date'
      );
      this.form.get(badDate)?.markAsTouched();
      return;
    }

    if (missing.length > 0 || invalidEstimates.length > 0) {
      // Mark touched so the template's *ngIf error messages light up too
      missing.forEach(f => this.form.get(f.control)?.markAsTouched());
      invalidEstimates.forEach(name => this.form.get(name)?.markAsTouched());

      const messages: string[] = [];
      if (missing.length > 0) {
        messages.push(`Please fill in: ${missing.map(f => f.label).join(', ')}`);
      }
      if (invalidEstimates.length > 0) {
        const labels = invalidEstimates.map(name =>
          name === 'estimatedVolume' ? 'Estimated Volume' : 'Estimated Dollars'
        );
        messages.push(`${labels.join(' and ')} must be greater than 0.`);
      }

      this.toastr.error(messages.join(' '), 'Invalid form');
      return;
    }

    const hasSpecsFile = this.attachments.specs.length > 0;

    if (!hasSpecsFile) {
      this.toastr.error('Please attach a file in the SPECS section', 'Error');
      return;
    }
    // Only enforce the "explain your changes" note requirement when this
    // is an update to a concept that's already fully submitted (not a
    // draft). The very submit that converts a draft into a main concept
    // (this.conceptId already exists from an earlier Save as Draft, but
    // isDraftConcept is still true) must NOT require a note — there's
    // nothing to "explain a change" against yet, since the concept was
    // never actually finalized before now.
        if (this.conceptId && !this.isDraftConcept) {
      const changedFields = this.getChangedWatchedFields();
      const attachmentsChanged = this.getAttachmentsChanged();
      // devNotes can no longer contain an unsaved entry — notes only land
      // in there after a successful save (see submitConcept()). The only
      // place a not-yet-saved note can be is the input itself.
      const hasNewNote = this.newNoteText.trim().length > 0;
      // conceptName is deliberately left out of watchedFields (see its
      // declaration) since editing it alone doesn't require a Development
      // Note. It still counts as a real edit for the "did anything change
      // at all" check below, though.
      const conceptNameChanged = !!this.form.get('conceptName')?.dirty;

      // Fields that differ from what's ACTUALLY persisted on the backend
      // right now (see lastPersistedFieldValues's doc comment) — catches
      // Client Approval having silently written a new value into a shared
      // field (estimatedVolume/estimatedDollars) and the user then editing
      // that field on the Development tab, INCLUDING clearing it back to
      // whatever the frozen originalFieldValues baseline still remembers.
      // getChangedWatchedFields() alone misses that case entirely, since
      // the cleared value matches the stale baseline even though it no
      // longer matches what's persisted.
      const fieldsDifferingFromPersisted = this.getFieldsDifferFromPersisted();

      // Union of both signals — a field counts as "changed" for BOTH the
      // no-op guard and the Development Note requirement if it differs
      // from the last Development-tab save OR from what's currently
      // persisted.
      const effectiveChangedFields = Array.from(
        new Set([...changedFields, ...fieldsDifferingFromPersisted])
      );

      // Nothing on this tab was actually touched — updating would just
      // write back exactly what's already saved. Block it instead of
      // hitting the backend for a no-op save with a false "updated
      // successfully" toast.
      if (!conceptNameChanged && effectiveChangedFields.length === 0 && !attachmentsChanged && !hasNewNote) {
        this.toastr.info('No changes to update.', 'Nothing to Save');
        return;
      }

      if ((effectiveChangedFields.length > 0 || attachmentsChanged) && !hasNewNote) {
        this.noteInputInvalid = true;
        // The note input lives inside the collapsible Development Notes
        // body — force it open before focusing, or the input (and the
        // red-outline/hint that explain why the submit was blocked)
        // would be invisible behind a collapsed card.
        this.notesExpanded = true;
        this.toastr.error(
          'You have changed tracked fields. Please add a Development Note explaining the changes before Updating.',
          'Development Note Required'
        );
        // Scroll note input into view so the user knows exactly what to fill
        setTimeout(() => document.querySelector<HTMLElement>('.note-input')?.focus());
        return;
      }
    }

    this.noteInputInvalid = false;

    await this.submitConcept(false);
  }

  /** Save as Draft — same upload pipeline as onSubmit, but skips most
   *  mandatory-field / SPECS-file checks since a draft is allowed to be
   *  incomplete. Still validates Concept Name, and — while creating a
   *  brand-new concept — Client Name, Master Concept Name, Review Type,
   *  and Claim Type, since those four can't be edited once the concept
   *  exists (see lockCoreFields()) and there'd be no way to fill them in
   *  later. Sends isDraft: 1 in the metadata so the backend can
   *  distinguish a draft save from a final submission. */
  async onSaveAsDraft(): Promise<void> {
    if (!this.canEdit) {
      this.toastr.error('You do not have permission to save drafts.', 'Access Denied');
      return;
    }
    // Safety net behind canSaveOrCreate/the [disabled] binding on the
    // button — Data Science Programmer never originates a new concept,
    // so Save-as-Draft stays blocked for that role even if triggered
    // some other way (e.g. Enter key) while it's still just a draft.
    if (this.canDSEdit && !this.developmentCompleted) {
      this.toastr.error('You do not have permission to save drafts.', 'Access Denied');
      return;
    }
    if (this.isRecordLocked && !(this.isProductionLocked && this.isManager)) {
      this.toastr.error(this.recordLockedMessage, 'Access Denied');
      return;
    }
    if (this.isAttachmentUploading) {
      this.toastr.error('Please wait for all files to finish uploading.', 'Upload in progress');
      return;
    }
    const conceptName = this.form.get('conceptName')?.value?.trim();
    if (!conceptName) {
      this.form.get('conceptName')?.markAsTouched();
      this.toastr.error('Concept Name is required even to save as draft', 'Error');
      return;
    }

    // Client Name / Master Concept Name / Review Type / Claim Type are
    // only required at creation time — once the concept exists they're
    // hidden (see template) and identified by concept_id alone, so this
    // check (like onSubmit()'s equivalent) only applies while creating a
    // brand-new concept. A draft is otherwise allowed to be incomplete,
    // but these four must still be filled in before the very first save.
    if (!this.conceptId) {
      const draftRequiredFields: { control: string; label: string }[] = [
        { control: 'clientName', label: 'Client Name' },
        { control: 'masterConceptName', label: 'Master Concept Name' },
        { control: 'reviewType', label: 'Review Type' },
        { control: 'claimType', label: 'Claim Type' },
      ];
      const missingDraftFields = draftRequiredFields.filter(f => {
        const value = this.form.get(f.control)?.value;
        return value === null || value === undefined || String(value).trim() === '';
      });
      if (missingDraftFields.length > 0) {
        missingDraftFields.forEach(f => this.form.get(f.control)?.markAsTouched());
        this.toastr.error(
          `Please fill in: ${missingDraftFields.map(f => f.label).join(', ')}`,
          'Required fields missing'
        );
        return;
      }
    }

    // Drafts skip most required-field checks, but a past-dated schedule
    // is a data-integrity issue regardless of draft status — block it
    // here too, since onSubmit()'s equivalent check never runs for a
    // draft save.
    const draftDateFields = ['qaSchedule', 'productionSchedule'];
    const badDraftDate = draftDateFields.find(f => this.form.get(f)?.hasError('pastDate'));
    if (badDraftDate) {
      this.toastr.error('QA Schedule and Production Schedule cannot be set in the past.', 'Invalid Date');
      this.form.get(badDraftDate)?.markAsTouched();
      return;
    }
    const badDraftDateFormat = draftDateFields.find(f => this.form.get(f)?.hasError('invalidDateRange'));
    if (badDraftDateFormat) {
      this.toastr.error('Please enter a valid date (MM/DD/YYYY).', 'Invalid Date');
      this.form.get(badDraftDateFormat)?.markAsTouched();
      return;
    }

    await this.submitConcept(true);
  }

  /** Shared upload pipeline used by both onSubmit and onSaveAsDraft.
   *  @param isDraft when true, sends isDraft: 1 in the metadata payload
   *  and skips upload entirely if there are no files at all (drafts may
   *  have nothing attached yet). */
    /** Shared upload pipeline used by both onSubmit and onSaveAsDraft.
   *  @param isDraft when true, sends isDraft: 1 in the metadata payload
   *  and skips upload entirely if there are no files at all (drafts may
   *  have nothing attached yet). */
  private async submitConcept(isDraft: boolean): Promise<void> {
    // Draft saves flip `savingDraft` only, so the Submit button (bound to
    // `loading`) doesn't visually react to a Save as Draft click — see the
    // comment on the `loading`/`savingDraft` declarations above.
    if (isDraft) {
      this.savingDraft = true;
    } else {
      this.loading = true;
    }
    // Captured up front — captureNewConceptId() sets this.conceptId /
    // this.isEditMode as soon as the create call returns, so checking
    // either of those *after* the submit completes can no longer tell us
    // whether this submit started out as a brand-new concept.
    const wasCreatingNew = !this.conceptId;

    // Captured up front, same reasoning as wasCreatingNew above — this is
    // the very submit that will (if it succeeds) convert a draft into a
    // main concept. Used both to skip the version-bump on the backend
    // (see the isDraftConversion metadata flag below) and to decide the
    // effective development status further down.
    const wasDraftConversion = !isDraft && this.isDraftConcept;

    // A brand-new concept has no development status chosen yet, so it
    // shows/behaves as "New" (see developmentStatusLabel's fallback).
    // Draft saves and later updates leave whatever status is already on
    // the concept alone. Only computed here — the form control itself
    // isn't updated until the request actually succeeds, below.
    const formDevStatus = this.form.get('developmentStatus')?.value || '';

    // Draft saves pass through exactly whatever's on the form — a draft is
    // allowed to have no status chosen yet, and that blank value is sent
    // as-is. A REAL submit (isDraft === false), whether that's a brand-new
    // concept's first Submit or a draft finally being converted into a main
    // concept, must never send a blank status to the backend — default it
    // to 'New' in that case. Neither case auto-advances past 'New' on its
    // own anymore: every concept starts life as 'New' and only moves
    // forward when a role with permission explicitly changes Development
    // Status.
    const effectiveDevStatus = isDraft ? formDevStatus : (formDevStatus || 'New');

    try {
      const user_id = Number(sessionStorage.getItem('userId'));
      const clientId = this.form.get('clientName')?.value;
      const masterId = this.form.get('masterConceptName')?.value;
      const reviewType = this.form.get('reviewType')?.value;
      const claimType = this.form.get('claimType')?.value;

      const clientName = this.clientOptions.find(c => c.client_id === clientId)?.client_id ?? '';
      const masterConceptName = this.masterConceptOptions.find(m => m.master_id === masterId)?.master_id ?? '';

      // The note input is never optimistically added to devNotes anymore —
      // it only gets saved (and shown in the list) right here, on Update.
      // So there's exactly one possible pending note: whatever's currently
      // typed. Already-persisted notes are never resent, since that would
      // make the backend insert duplicate rows for them on every update.
      const pendingNote = this.newNoteText.trim();
      const developmentNotes = pendingNote ? [pendingNote] : [];

      const metadata: any = {
        isDraft: isDraft ? 1 : 0,
        // Tells the backend whether this submit is finalizing a concept
        // that was previously a draft — i.e. its very first "real" submit
        // — versus editing an already-finalized concept. The backend
        // should only bump CurrentConceptId's version suffix
        // (…_D001 -> …_D002) in the latter case; converting a draft into
        // a main concept for the first time must not count as a version
        // bump.
        isDraftConversion: wasDraftConversion,
        // Identifies which concept to update. Empty/omitted on the very
        // first submit (no concept exists yet) — the backend treats that
        // as a create. Every submit after that is an update against this id.
        concept_id: this.conceptId || undefined,
        conceptName: this.form.get('conceptName')?.value,
        InternalConceptDescription: this.form.get('Internalconceptdescription')?.value,
        developmentStatus: effectiveDevStatus,
        priority: this.form.get('priority')?.value,
        haloNumber: this.numericOrNull(this.form.get('haloNumber')?.value),
        developmentNotes,
        estimatedVolume: this.numericOrNull(this.form.get('estimatedVolume')?.value),
        estimatedDollars: this.numericOrNull(this.form.get('estimatedDollars')?.value),
        confidenceScore: this.form.get('confidenceScore.value')?.value,
        previousReportId: this.form.get('previousReportId')?.value,
        qaSchedule: this.form.get('qaSchedule')?.value,
        productionSchedule: this.form.get('productionSchedule')?.value,
        ideationRequestor: this.form.get('ideationRequestor')?.value,
        dataScienceProgrammer: this.form.get('dataScienceProgrammer')?.value,
        // A real (non-draft) submit IS the act of completing the
        // Development tab, so it must be sent as 1 in this very request —
        // not read off `this.developmentCompleted`, which is still 0 at
        // this point on the very first ("Create Concept") submit and only
        // gets flipped to 1 further down *after* this request is built.
        DevelopmentCompleted: isDraft ? this.developmentCompleted : 1,
        ClientApprovalCompleted: this.clientApprovalCompleted,
        SupportingDocumentsCompleted: this.supportingDocumentsCompleted,
      };

      // Client Name, Master Concept Name, Review Type, and Claim Type are
      // locked (view-only, see lockCoreFields) once the concept exists —
      // only send them on the very first create. On every later update
      // the backend already has them tied to concept_id, so they're left
      // out of the payload entirely instead of being resent unchanged.
      if (!this.conceptId) {
        metadata.clientName = clientName;
        metadata.masterConceptName = masterConceptName;
        metadata.reviewType = reviewType;
        metadata.claimType = claimType;
      }
      const metadataJson = JSON.stringify(metadata);
      const categories: AttachCategory[] = ['specs', 'table', 'other'];

      // Restored attachments are zero-byte placeholders (progress 100,
      // size 0) — they count for display purposes, but there's nothing
      // new in them to upload. Only entries with real bytes belong here;
      // everything else falls through to the metadata-only branch below,
      // otherwise a text-only update (no new files touched) would enter
      // the upload loop, skip every placeholder via `continue`, and never
      // actually call the API at all.
      const filesToUpload: { cat: AttachCategory; entry: AttachFile }[] = [];
      for (const cat of categories) {
        for (const entry of this.attachments[cat]) {
          if (entry.file.size === 0 && entry.progress === 100) continue;
          filesToUpload.push({ cat, entry });
        }
      }

      // Single combined request: metadata + every new attachment from all
      // three categories (specs/table/other) go out together in ONE
      // FormData/API call, instead of one call per file (or per file per
      // chunk) like before. Files are told apart on the backend side by
      // the parallel 'categories' array, matched by position to 'files'.
      const formData = new FormData();
      formData.append('user_id', user_id.toString());
      formData.append('metadata', metadataJson);
      formData.append('concept_id', this.conceptId || '');

      filesToUpload.forEach(({ cat, entry }) => {
        formData.append('files', entry.file, entry.file.name);
        formData.append('categories', cat);
        formData.append('file_names', entry.file.name);
        formData.append('file_sizes', entry.file.size.toString());
      });

      const res: any = await this.service.createconcept(formData).toPromise();
      if (res?.status === 'error' || res?.detail) {
        this.toastr.error(res?.detail || res?.message || 'Something went wrong', 'Error');
        return;
      }
      this.captureNewConceptId(res);

      // Backend's explicit signal for a Development -> Production
      // transition (see moved_to_production in create_concept.py) — used
      // below to give this milestone its own toast instead of the
      // generic "updated"/"version created" ones, and to suppress those
      // so a single save never fires two overlapping toasts.
      const backendStatusMessage: string | null = res?.status_message ?? null;

      // Backend's explicit signal for a Production -> Superseded
      // transition: the current concept just went fully read-only and a
      // new development cycle was created alongside it. Captured now so
      // it can drive the switch-confirmation popup once the current
      // record's own save feedback (toast + soft-reload below) has run.
      const supersededNewConceptId: string | null = res?.new_concept_id ?? null;

      if (!isDraft && effectiveDevStatus !== formDevStatus) {
        // Covers a brand-new concept's first submit, and the
        // draft-with-no-status → main conversion: the backend was just
        // sent 'New' even though the form itself still shows blank —
        // sync the form/status dropdown to match what was actually
        // persisted.
        this.form.get('developmentStatus')?.setValue(effectiveDevStatus, { emitEvent: false });
        this.updateEstimatedFieldsValidators(effectiveDevStatus);
        this.refreshAllowedStatuses(effectiveDevStatus);
      }
      // Surface the version bump to the user — the backend is expected to
      // report version_updated: false on the draft -> main conversion
      // submit (see the isDraftConversion flag sent above), so this toast
      // naturally stops firing for that case without any extra guard here.
      // Also skipped when movedToProduction, which gets its own more
      // specific toast further down instead.
      if (res?.version_updated && res?.current_concept_id && !backendStatusMessage) {
        this.toastr.info(
          `New version created: ${res.current_concept_id}`,
          'Version Updated'
        );
      }

      // Mark every newly uploaded attachment complete now that the single
      // request has succeeded (no per-chunk progress to track anymore).
      filesToUpload.forEach(({ entry }) => { entry.progress = 100; });

      // A SPECS replace staged via confirmSpecsReplace() only deletes the
      // old file's backend record now — after the replacement has just
      // been persisted above. Deleting it any earlier (e.g. at confirm
      // time) could leave SPECS with zero files server-side if this save
      // never completed. Best-effort: if this cleanup call fails, the new
      // file is already safely saved, so just warn instead of rolling
      // anything back.
      if (this.pendingSpecsDeletion) {
        const staleAttachmentId = this.pendingSpecsDeletion.attachmentId;
        this.pendingSpecsDeletion = null;
        if (staleAttachmentId) {
          try {
            const cleanupForm = new FormData();
            cleanupForm.append('concept_id', this.conceptId);
            cleanupForm.append('attachment_id', staleAttachmentId.toString());
            cleanupForm.append('category', 'specs');
            cleanupForm.append('user_id', user_id.toString());
            await this.service.deleteattachment(cleanupForm).toPromise();
          } catch (err: any) {
            const msg = err?.error?.detail || err?.error?.message || err?.message
              || 'The replaced SPECS file was saved, but the old file could not be removed from storage.';
            this.toastr.error(msg, 'Cleanup Warning');
          }
        }
      }

      if (isDraft) {
        this.toastr.success('Draft saved successfully!', 'Success');
      } else if (backendStatusMessage) {
        // Distinct, explicit confirmation for this milestone — checked
        // before the create/update branch below so it always wins, even
        // in the unlikely case both are somehow true for the same save.
        this.toastr.success(backendStatusMessage, 'Success');
      } else {
        this.toastr.success(
          (wasCreatingNew || wasDraftConversion)
            ? 'Concept Created successfully!'
            : 'Concept updated successfully!',
          'Success'
        );
      }
      if (!isDraft) {
        this.developmentCompleted = 1;
        // this.isDraftConcept = false;
      }

      // The pending note (if any), plus every note already in devNotes,
      // just went out inside `metadata` above and is now saved on the
      // backend. Mark everything persisted so the next submit only sends
      // genuinely new notes — otherwise this same list gets re-inserted
      // as duplicates on the next Update.
      this.devNotes = this.devNotes.map(n => ({ ...n, persisted: true }));
      if (pendingNote) {
        this.devNotes = [...this.devNotes, {
          id: `n${Date.now()}`, author: this.currentUserName, initials: this.currentUserInitial,
          avatarBg: '#6366f1', time: new Date().toISOString(), text: pendingNote,
          RoleName: sessionStorage.getItem('roleName') ?? '',
          persisted: true
        }];
        this.newNoteText = '';
        this.resetNoteInputHeight();
        this.scrollNotesToBottom();
      }

      // Refresh the left panel so the submitted/updated concept shows
      // up there immediately, without needing a full page reload. This is
      // a quiet background re-sync — see loadLatestUpdates()'s
      // `background` param — so it doesn't blink the whole list out
      // behind a spinner for a save the user already saw succeed via the
      // toast above.
      this.loadLatestUpdates(true);

      // Only when this submit (a) was a genuine final submit, not a draft
      // save, and (b) started out with no conceptId at all — i.e. this
      // really was "create a new concept". Land the user ON the concept
      // they just created instead of resetting back to a blank form —
      // captureNewConceptId() above already set this.conceptId to the
      // newly-assigned anchor id, so navigate straight to it. This
      // triggers the routeSub in ngOnInit (paramMap changes from no id ->
      // this concept's id), which runs loadConcept() and fully re-patches
      // the page from exactly what the server just persisted.
      if (!isDraft && wasCreatingNew && this.conceptId) {
        this.router.navigate(['/concept-create', this.conceptId]);
      } else if (this.conceptId) {
        // Soft-reload so every field reflects exactly what the server
        // just persisted — no full page reload, no tab switch. This is
        // the one soft refresh allowed to re-baseline the Development
        // Note "changed fields" tracking (see the note on
        // fetchAndApplyConcept()), since this save is the one that just
        // legitimately satisfied (or didn't need) that requirement.
        this.refreshConceptData(this.conceptId, /* resnapshotDevelopmentFields */ true);
      }

      // This save is what just superseded the concept — it's now fully
      // read-only (see isProductionLocked/isReadOnly), so instead of
      // leaving the Manager sitting on a locked record, ask whether to
      // jump straight into the new development cycle the backend just
      // spun up. Deliberately opt-in rather than automatic: they may
      // still want to review the now-superseded record first.
      if (supersededNewConceptId) {
        this.pendingSupersededConceptId = supersededNewConceptId;
        this.showSupersededSwitchModal = true;
      }
    } catch (err: any) {
      const errorMsg =
        err?.error?.detail ||
        err?.error?.message ||
        err?.message ||
        'Upload failed. Please try again.';
      this.toastr.error(this.friendlyErrorMessage(errorMsg), 'Error');
    } finally {
      if (isDraft) {
        this.savingDraft = false;
      } else {
        this.loading = false;
      }
    }
  }

  /** When creating a brand-new concept, the very first save call is what
   *  assigns its ID on the backend. Capture it here so the *next* save in
   *  the same session (e.g. switching tabs and hitting Submit again)
   *  updates that same concept instead of creating a duplicate.
   *  NOTE: adjust the field names below to match your actual API response shape. */
  // private captureNewConceptId(res: any): void {
  //   // res.concept_id is the stable anchor and should never change once set
  //   // — only capture it the first time (on genuine create). res.current_concept_id
  //   // is the version/display id and DOES change on every update, so it's
  //   // always safe (and necessary) to refresh displayConceptId here.
  //   const anchorId = res?.concept_id ?? res?.ConceptId ?? res?.data?.ConceptId ?? res?.concept?.ConceptId;
  //   const currentId = res?.current_concept_id ?? res?.CurrentConceptId ?? anchorId;

  //   if (!this.conceptId && anchorId) {
  //     this.conceptId = anchorId;
  //     this.isEditMode = true;
  //     this.lockCoreFields();
  //   }
  //   if (currentId) {
  //     this.displayConceptId = currentId;
  //   }
  // }
  private captureNewConceptId(res: any): void {
    const anchorId = res?.concept_id ?? res?.ConceptId ?? res?.data?.ConceptId ?? res?.concept?.ConceptId;
    const currentId = res?.current_concept_id ?? res?.CurrentConceptId ?? anchorId;

    const wasEmpty = !this.conceptId;
    if (anchorId) {
      this.conceptId = anchorId;       // always resync, not just when empty
    }
    if (wasEmpty && anchorId) {
      this.isEditMode = true;
      this.lockCoreFields();
    }
    if (currentId) {
      this.displayConceptId = currentId;
    }
  }
  /** Client Name, Master Concept Name, Review Type, and Claim Type are
   *  set once at creation and shouldn't change afterward — lock them
   *  (view-only) as soon as the concept exists, whether that's a
   *  just-created concept in this session or one restored from a saved
   *  record. submitConcept() separately checks isEditMode/conceptId to
   *  stop resending these in the update payload — see there. */
  private lockCoreFields(): void {
    const fieldsToLock = ['clientName', 'masterConceptName', 'reviewType', 'claimType'];

    // Ideation Requestor may only be reassigned by a Manager once the
    // concept has been created. Every other role — including the
    // Ideation Requestor who originated it and QA (the DS Programmer is
    // already locked out of this field separately, in
    // applyRoleRestrictions()) — sees it locked read-only from here on.
    // Managers are excluded from the lock so they retain the ability to
    // reassign it.
    if (this.currentUserRoleId !== ConceptCreateComponent.ROLE_MANAGER) {
      fieldsToLock.push('ideationRequestor');
    }

    fieldsToLock.forEach(name => {
      this.form.get(name)?.disable({ emitEvent: false });
    });
  }

  // ── Add new concept ───────────────────────────────────────────────────
  onAddNewConcept(): void {
    if (!this.canCreateConcept) {
      this.toastr.error('You do not have permission to create a new concept.', 'Access Denied');
      return;
    }
    // Navigating to the same '/concept-create' URL we're already on (e.g.
    // mid-draft — submitConcept() never pushes the new conceptId into the
    // URL) is a no-op for the router: no NavigationEnd, no paramMap
    // emission, so the routeSub in ngOnInit() never calls
    // resetToNewConcept() and the button appears broken. Reset directly
    // instead of depending on that route change firing.
    this.resetToNewConcept();

    if (this.router.url !== '/concept-create') {
      this.router.navigate(['/concept-create']);
    }
  }

  // ── Copy concept ──────────────────────────────────────────────────────
  /** Entry point for the Copy button — validates, then gates the actual
   *  copy behind the shared "Are you sure?" confirmation popup (see
   *  openConfirmModal() above) so a stray/misclick doesn't immediately
   *  spawn a new concept. performCopyConcept() only runs if the user
   *  confirms. */
  onCopyConcept(): void {
    if (!this.canCreateConcept) {
      this.toastr.error('You do not have permission to create a new concept.', 'Access Denied');
      return;
    }
    // Nothing loaded yet (blank creation form) — nothing to copy.
    if (!this.conceptId || this.copyingConcept) return;

    this.openConfirmModal(
      `This will create a new concept pre-filled from ${this.displayConceptId || this.conceptId}. You'll still need to review and complete it before it's saved.`,
      () => this.performCopyConcept(),
      'Copy'
    );
  }

  /** Copies the currently-loaded concept into a brand-new, unsaved one.
   *  Hits POST copy-concept with this concept's id, then lands the user on
   *  the blank "new concept" form pre-filled from the response — same
   *  destination as "+ Add New Concept", just not blank. The user still
   *  has to fill in Client/Master Concept/Review/Claim Type (the endpoint
   *  doesn't return them) and hit Create to actually persist it; nothing
   *  is saved by the copy call itself. Only ever called after the user
   *  confirms via onCopyConcept()'s popup. */
  private performCopyConcept(): void {
    this.copyingConcept = true;
    this.service.copyConcept(this.conceptId).subscribe({
      next: (res: any) => {
        this.copyingConcept = false;
        if (!res?.success) {
          this.toastr.error('Failed to copy concept.', 'Copy Failed');
          return;
        }
        this.applyCopiedConcept(res);
        this.toastr.success('Review the details and click Create to save it.', 'Concept Copied');
      },
      error: (err) => {
        this.copyingConcept = false;
        console.error('Failed to copy concept:', err);
        this.toastr.error('Failed to copy concept.', 'Copy Failed');
      }
    });
  }

  /** Kicks off landing the copy-concept response onto a fresh "new
   *  concept" form.
   *
   *  This can't just reset -> patch -> navigate (the previous approach):
   *  navigating from /concept-create/:id to /concept-create is a
   *  different route config, not a changed :id param on the same one, so
   *  Angular's default reuse strategy destroys this component and
   *  constructs a brand-new instance for the destination route rather
   *  than reusing this one. Patching `this.form` before/after calling
   *  router.navigate() was patching THIS (outgoing, soon-to-be-destroyed)
   *  instance either way — the new instance that actually renders has its
   *  own fresh `this.form` that was never touched, which is why the copied
   *  data silently never appeared.
   *
   *  Fix: pass the metadata through router.navigate()'s `state` extra
   *  instead of touching `this.form` at all here. That state travels with
   *  the navigation itself (not with this component instance), so the
   *  new instance's constructor can pick it up via
   *  router.getCurrentNavigation() regardless of whether it's a fresh
   *  instance or a reused one, and its routeSub applies it once the blank
   *  form actually exists to receive it — see the constructor and
   *  routeSub in ngOnInit, and applyCopiedMetadataToForm() below. */
  private applyCopiedConcept(res: any): void {
    const metadata = res.metadata ?? {};

    if (this.router.url !== '/concept-create') {
      this.router.navigate(['/concept-create'], { state: { copiedConceptMetadata: metadata } });
    } else {
      // Already sitting on the blank creation route (shouldn't normally
      // happen — the Copy button only shows once a real concept with an
      // id is loaded — but handled for safety). The route isn't changing,
      // so neither the constructor's getCurrentNavigation() read nor
      // routeSub will fire for it; reset and apply directly instead.
      this.resetToNewConcept();
      this.applyCopiedMetadataToForm(metadata);
    }
  }

  /** Patches a copy-concept response's metadata onto the current
   *  `this.form` — called only once that form is confirmed to belong to
   *  the blank "new concept" page it's meant for (see applyCopiedConcept()
   *  and the routeSub in ngOnInit). */
  private applyCopiedMetadataToForm(m: any): void {
    // conceptName arrives as a single-item array (["Mesh-10"]) from this
    // endpoint rather than a plain string like every other field here —
    // unwrap it defensively either way in case that ever changes.
    const conceptName = Array.isArray(m.conceptName)
      ? (m.conceptName[0] ?? '')
      : (m.conceptName ?? '');

    this.form.patchValue({
      conceptName,
      priority: m.priority ?? '',
      haloNumber: m.haloNumber ?? '',
      Internalconceptdescription: m.InternalConceptDescription ?? '',
      developmentStatus: m.developmentStatus ?? '',
      estimatedVolume: m.estimatedVolume ?? '',
      estimatedDollars: m.estimatedDollars ?? '',
      previousReportId: m.previousReportId ?? '',
      qaSchedule: m.qaSchedule ? m.qaSchedule.split('T')[0] : '',
      productionSchedule: m.productionSchedule ? m.productionSchedule.split('T')[0] : '',
      ideationRequestor: m.ideationRequestor ?? '',
      dataScienceProgrammer: m.dataScienceProgrammer ?? '',
      // clientName / masterConceptName / reviewType / claimType are
      // deliberately left blank — copy-concept doesn't return them, and
      // they're required fields the user must (re)confirm for the new
      // concept, same as any fresh creation.
    });

    if (m.confidenceScore !== undefined && m.confidenceScore !== null && m.confidenceScore !== '') {
      this.form.get('confidenceScore.value')?.setValue(this.confidenceScoreToBucket(m.confidenceScore));
    }

    this.updateEstimatedFieldsValidators(m.developmentStatus ?? 'New');

    // A fresh copy has nothing saved yet, so nothing here should read as
    // an unsaved edit.
    this.form.markAsPristine();
    this.form.markAsUntouched();
  }

  /** The Confidence Score control only ever stores 'low' | 'medium' | 'high'
   *  (see setConfidenceScore()/the toggle buttons) — but copy-concept sends
   *  a raw 0–100 score ("95") instead of one of those words. Buckets it the
   *  same way the UI's three-way toggle implies. NOTE: the >=85 / >=50
   *  thresholds below are a reasonable guess, not a confirmed business
   *  rule — adjust if the backend has an official cutoff for this. */
  private confidenceScoreToBucket(raw: string | number): 'low' | 'medium' | 'high' {
    const word = String(raw).trim().toLowerCase();
    if (word === 'low' || word === 'medium' || word === 'high') return word;

    const n = Number(raw);
    if (!isNaN(n)) {
      if (n >= 85) return 'high';
      if (n >= 50) return 'medium';
      return 'low';
    }
    return 'low';
  }

  private resetToNewConcept(): void {
    // Same reasoning as the routeSub fix above — resetToNewConcept() can
    // also be called directly (onAddNewConcept, when already sitting on
    // /concept-create with no id, where the route doesn't actually change
    // and routeSub never fires) — so it needs its own clear, not just a
    // reliance on the routeSub subscriber.
    this.dismissDraftLockBanner();
    this.isEditMode = false;
    this.isDraftConcept = false;
    this.buildForm();
    this.conceptId = '';
    this.displayConceptId = '';
    this.savedConceptName = '';
    this.createdDate = new Date();
    this.updatedDate = new Date();
    this.uploadedFiles = [];
    this.owners = [];
    this.devNotes = [];
    this.newNoteText = '';
    this.resetNoteInputHeight();
    // See the matching comment in loadConcept() — a blocked Update on the
    // previous concept can leave this true, and it must not carry over
    // onto a brand-new, untouched concept.
    this.noteInputInvalid = false;
    // Same reasoning — a forced-open card from a blocked Update on the
    // previous concept shouldn't carry over onto a brand-new one either.
    this.notesExpanded = false;
    // A fresh/new concept has no assigned Ideation Requestor / DS Programmer
    // of its own — clear the previous concept's stashed fallback so it
    // doesn't leak into this blank form (see ensureAssignedUsersVisible()).
    this.lastConceptRequestor = null;
    this.lastConceptProgrammer = null;
    this.attachments = { specs: [], table: [], other: [], approval: [] };
    this.originalAttachmentIds = { specs: new Set(), table: new Set(), other: new Set(), approval: new Set() };
    // Clear the previous concept's audit trail too — a brand-new concept
    // has no history of its own, and without this the "Activity history"
    // card keeps showing the last-loaded concept's entries (and a stale
    // selectedActivity popup could still be open from it).
    this.activityHistory = [];
    this.activityHistoryLoading = false;
    this.showFullActivityHistory = false;
    this.selectedActivity = null;
    this.showActivityLegend = false;
    // Same reasoning — the previous concept's dot-color legend (built from
    // ITS dot_color_map response) has no business showing on a brand-new
    // concept that hasn't loaded any history of its own yet.
    this.activityDotLegend = [];
    this.activeTab = 'development';
    this.developmentCompleted = 0;
    this.clientApprovalCompleted = 0;
    this.supportingDocumentsCompleted = 0;
    // Clear the "submit was attempted" flags too — otherwise the brand-new
    // form's blank (and therefore required-invalid) estimatedVolume /
    // estimatedDollars controls immediately show red validation errors on
    // this fresh page, even though the user hasn't touched them here. These
    // flags were left true by the submit that just succeeded and produced
    // this new-concept form in the first place; they must not carry over.
    this.developmentSubmitAttempted = false;
    this.approvalSubmitAttempted = false;
    // Brand-new concept — no prior DocIndex values exist for it, so the
    // allocator can safely restart from 0.
    this.nextDocIndex = 0;
    this.supportingDocs = [
      this.blankDoc(),
      this.blankDoc(),
      this.blankDoc()
    ];
    this.applyRoleRestrictions();
    this.refreshAllowedStatuses('New');
    // loadMasterData() only runs once per page load (see ngOnInit), so on a
    // create -> save -> create-again cycle the options list is already
    // populated by the time we land back here — try the auto-fill
    // immediately. If the options aren't loaded yet (very first load of a
    // brand-new concept), loadMasterData()'s callback retries this once
    // they arrive.
    this.prefillIdeationRequestor();
    this.scrollTabContentToTop();
  }

  /** Auto-fills "Ideation Requestor" with the logged-in user's own entry
   *  when creating a brand-new concept — but ONLY for roles who can
   *  actually be an Ideation Requestor: Ideation Requestor, QA, and
   *  Manager (the same set as canFullEdit). Every other role (e.g. Data
   *  Science Programmer) leaves the field on "Select" and picks someone
   *  from the dropdown manually. Skipped entirely in edit mode, where the
   *  field is patched from the saved concept's own data instead — see
   *  patchForm(). */
  private prefillIdeationRequestor(): void {
    if (this.isEditMode) return;
    if (!this.canFullEdit) return;

    const control = this.form.get('ideationRequestor');
    if (!control || control.value) return; // already has a value — don't clobber it

    const userId = Number(sessionStorage.getItem('userId'));
    if (!userId) return;

    const match = this.ideationRequestorOptions.find(u => u.id === userId);
    if (match) {
      control.setValue(match.id, { emitEvent: false });
    }
  }

  // ── Ownership ─────────────────────────────────────────────────────────
  addOwnership(): void {
    this.showOwnerDropdown = !this.showOwnerDropdown;
    this.ownerSearch = '';
  }

  selectAndAddOwner(c: OwnerCandidate): void {
    this.owners = [...this.owners, {
      id: c.id, name: c.name, department: c.department,
      initials: c.initials, avatarColor: c.avatarColor
    }];
    this.showOwnerDropdown = false;
    this.ownerSearch = '';
  }

  cancelOwnerDropdown(): void {
    this.showOwnerDropdown = false;
    this.ownerSearch = '';
  }

  removeOwner(id: string): void {
    this.owners = this.owners.filter(o => o.id !== id);
  }

  // ── Dev Notes ─────────────────────────────────────────────────────────
  // There's no "add to list" action here anymore — the typed text in
  // newNoteText just sits in the input until Update is clicked. That's
  // picked up as `pendingNote` in submitConcept() and only added to
  // devNotes (so it becomes visible) once the save actually succeeds.

  private scrollNotesToBottom(): void {
    setTimeout(() => {
      const el = this.notesListRef?.nativeElement;
      if (el) el.scrollTop = el.scrollHeight;
    });
  }

  // ── Attachments ───────────────────────────────────────────────────────
  triggerAttachUpload(cat: AttachCategory): void {
    if (this.blockIfCannotManage(cat)) return;
    const map: Record<AttachCategory, ElementRef<HTMLInputElement>> = {
      specs: this.specsInputRef,
      table: this.tableInputRef,
      other: this.otherInputRef,
      approval: this.approvalInputRef
    };
    map[cat]?.nativeElement.click();
  }

  private readonly allowedExtensions = ['pdf', 'doc', 'docx', 'xls', 'xlsx'];

  private isValidFile(file: File): boolean {
    const extension = file.name.split('.').pop()?.toLowerCase() || '';
    return this.allowedExtensions.includes(extension);
  }

  onAttachSelected(event: Event, cat: AttachCategory): void {
    if (this.blockIfCannotManage(cat)) return;

    const input = event.target as HTMLInputElement;
    if (!input.files || input.files.length === 0) return;

    // SPECS is single-file only — route through the replace-confirmation
    // flow instead of the normal multi-file loop below.
    if (cat === 'specs') {
      if (input.files.length > 1) {
        this.toastr.error(
          'Only one file can be attached under SPECS. Please select a single file.',
          'Single File Only'
        );
      }
      this.handleSpecsFile(input.files[0]);
      input.value = '';
      return;
    }

    const duplicateNames: string[] = [];
    const invalidFiles: string[] = [];

    Array.from(input.files).forEach(file => {

      // Validate file type
      if (!this.isValidFile(file)) {
        invalidFiles.push(file.name);
        return;
      }

      // Check duplicate filenames
      const isDuplicate = this.attachments[cat].some(
        existing =>
          existing.name.trim().toLowerCase() ===
          file.name.trim().toLowerCase()
      );

      if (isDuplicate) {
        duplicateNames.push(file.name);
        return;
      }

      const entry: AttachFile = {
        id: this.generateAttachId(),
        name: file.name,
        size: file.size,
        progress: 0,
        file
      };

      this.attachments[cat] = [...this.attachments[cat], entry];
      this.simulateUpload(cat, entry.id);
    });

    // Show invalid file message
    if (invalidFiles.length > 0) {
      this.toastr.error(
        `${invalidFiles.map(n => `"${n}"`).join(', ')} ${invalidFiles.length > 1 ? 'are' : 'is'
        } not a supported file type. Only PDF, Word (.doc/.docx), and Excel (.xls/.xlsx) files are allowed.`,
        'Invalid File Type'
      );
    }

    // Show duplicate file message
    if (duplicateNames.length > 0) {
      this.toastr.error(
        `${duplicateNames.map(n => `"${n}"`).join(', ')} ${duplicateNames.length > 1 ? 'are' : 'is'
        } already attached in this section.`,
        'Duplicate File'
      );
    }

    input.value = '';
  }

  onAttachDragOver(event: DragEvent): void {
    event.preventDefault(); event.stopPropagation();
  }

  onAttachDrop(event: DragEvent, cat: AttachCategory): void {
    event.preventDefault(); event.stopPropagation();
    if (this.blockIfCannotManage(cat)) return;
    const files = event.dataTransfer?.files;
    if (!files || files.length === 0) return;

    // SPECS is single-file only — same replace-confirmation flow as
    // onAttachSelected() above, drag/drop is just another entry point.
    if (cat === 'specs') {
      if (files.length > 1) {
        this.toastr.error(
          'Only one file can be attached under SPECS. Please drop a single file.',
          'Single File Only'
        );
      }
      this.handleSpecsFile(files[0]);
      return;
    }

    // Same duplicate-filename guard as onAttachSelected() above — drag/drop
    // is just another entry point for adding a file to this same array.
    const duplicateNames: string[] = [];
    Array.from(files).forEach(file => {
      const isDuplicate = this.attachments[cat].some(
        existing => existing.name.trim().toLowerCase() === file.name.trim().toLowerCase()
      );
      if (isDuplicate) {
        duplicateNames.push(file.name);
        return;
      }
      const entry: AttachFile = { id: this.generateAttachId(), name: file.name, size: file.size, progress: 0, file };
      this.attachments[cat] = [...this.attachments[cat], entry];
      this.simulateUpload(cat, entry.id);
    });
    if (duplicateNames.length > 0) {
      this.toastr.error(
        `${duplicateNames.map(n => `"${n}"`).join(', ')} ${duplicateNames.length > 1 ? 'are' : 'is'} already attached in this section.`,
        'Duplicate File'
      );
    }
  }

  /** Validates a newly picked/dropped SPECS file and either attaches it
   *  directly (no existing SPECS file yet) or stages it behind the
   *  replace-confirmation modal (an existing SPECS file would otherwise
   *  be silently swapped out from under the user). */
  private handleSpecsFile(file: File): void {
    if (!this.isValidFile(file)) {
      this.toastr.error(
        `"${file.name}" is not a supported file type. Only PDF, Word (.doc/.docx), and Excel (.xls/.xlsx) files are allowed.`,
        'Invalid File Type'
      );
      return;
    }

    const existing = this.attachments.specs[0];

    if (!existing) {
      this.addSingleAttachFile('specs', file);
      return;
    }

    if (existing.name.trim().toLowerCase() === file.name.trim().toLowerCase()) {
      this.toastr.error(`"${file.name}" is already attached in this section.`, 'Duplicate File');
      return;
    }

    this.pendingSpecsFile = file;
    this.pendingSpecsExisting = existing;
    this.showSpecsReplaceModal = true;
  }

  /** User confirmed the popup — deactivate the existing SPECS file locally
   *  and attach the newly selected file in its place. The old file's
   *  backend record is deliberately NOT deleted here — see
   *  pendingSpecsDeletion's comment above for why that has to wait until
   *  submitConcept() actually persists the replacement. */
  confirmSpecsReplace(): void {
    const newFile = this.pendingSpecsFile;
    const existing = this.pendingSpecsExisting;
    this.showSpecsReplaceModal = false;
    this.pendingSpecsFile = null;
    this.pendingSpecsExisting = null;

    if (!newFile || !existing) return;

    // Only a persisted file needs a backend deletion at all. If a
    // deletion is already staged from an earlier replace in this same
    // editing session, keep tracking that original persisted file rather
    // than overwriting it with an unsaved intermediate one that has
    // nothing to delete on the server.
    if (existing.attachmentId && !this.pendingSpecsDeletion) {
      this.pendingSpecsDeletion = existing;
    }

    this.attachments.specs = this.attachments.specs.filter(x => x !== existing);
    this.addSingleAttachFile('specs', newFile);
  }

  /** User dismissed the popup — the newly picked file is discarded and the
   *  existing SPECS file stays untouched. */
  cancelSpecsReplace(): void {
    this.showSpecsReplaceModal = false;
    this.pendingSpecsFile = null;
    this.pendingSpecsExisting = null;
  }

  /** User confirmed switching into the new development cycle that was
   *  created when this concept was superseded — navigate there. Routing
   *  to a different concept id fires the routeSub in ngOnInit (paramMap
   *  change) which runs loadConcept() and fully patches the page for the
   *  new record, same as the create-a-new-concept navigation above. */
  confirmSwitchToSupersededConcept(): void {
    const targetId = this.pendingSupersededConceptId;
    this.dismissSupersededSwitchModal();
    if (targetId) {
      this.router.navigate(['/concept-create', targetId]);
    }
  }

  /** User chose to stay put — the current concept was already refreshed
   *  as read-only by submitConcept()'s soft-reload, so there's nothing
   *  else to undo here, just close the popup. They can switch later from
   *  the concept's own history/activity trail. */
  dismissSupersededSwitchModal(): void {
    this.showSupersededSwitchModal = false;
    this.pendingSupersededConceptId = null;
  }

  /** Shared tail end of "add one validated file to a category" — used both
   *  for the first SPECS file and for the replacement after confirmation. */
  private addSingleAttachFile(cat: AttachCategory, file: File): void {
    const entry: AttachFile = {
      id: this.generateAttachId(),
      name: file.name,
      size: file.size,
      progress: 0,
      file
    };
    this.attachments[cat] = [...this.attachments[cat], entry];
    this.simulateUpload(cat, entry.id);
  }

  /** A file just picked this session (never submitted) has no backend
   *  record yet — drop it locally, nothing to call. A file restored from
   *  a saved concept has a real AttachmentId, so it has to be deleted on
   *  the backend first; the card is only removed from the UI once that
   *  call succeeds, so a failed delete doesn't silently desync the UI
   *  from what's actually still stored server-side. */
  async removeAttachment(cat: AttachCategory, f: AttachFile): Promise<void> {
    if (this.blockIfCannotManage(cat)) return;

    // SPECS must always have at least one file on an existing concept —
    // block removing the last one instead of letting it disappear (either
    // locally, for a not-yet-submitted file, or via the delete API, for a
    // persisted one) and leave the concept spec-less. This must run BEFORE
    // the attachmentId branch below: an unsubmitted file (no attachmentId)
    // would otherwise skip this check entirely by taking the early-return
    // path, letting the user delete every SPECS file — persisted ones
    // included — one at a time without ever tripping the guard. Checking
    // "how many remain after removing f" (not just the array's raw length)
    // also means a stray unsubmitted file can't pad the count and mask
    // that the last real file is about to go.
    if (cat === 'specs') {
      const remaining = this.attachments.specs.filter(x => x !== f).length;
      if (remaining === 0) {
        // If a replace is currently staged (see confirmSpecsReplace()),
        // the sole remaining SPECS file is that unsaved replacement, and
        // the original persisted file was never actually deleted from the
        // backend. Removing it just means "never mind" — restore the
        // original instead of leaving SPECS empty or blocking the removal.
        if (this.pendingSpecsDeletion) {
          this.attachments.specs = [this.pendingSpecsDeletion];
          this.pendingSpecsDeletion = null;
          return;
        }
        this.toastr.error(
          'At least one SPECS file is required. Upload a replacement before removing this one.',
          'Cannot Remove'
        );
        return;
      }
    }

    if (!f.attachmentId) {
      this.attachments[cat] = this.attachments[cat].filter(x => x !== f);
      return;
    }

    const user_id = Number(sessionStorage.getItem('userId'));
    try {
      const formData = new FormData();
      formData.append('concept_id', this.conceptId);
      formData.append('attachment_id', f.attachmentId.toString());
      formData.append('category', cat);
      formData.append('user_id', user_id.toString());

      await this.service.deleteattachment(formData).toPromise();
      this.attachments[cat] = this.attachments[cat].filter(x => x !== f);
      this.toastr.success('File deleted successfully!', 'Success');
      this.loadLatestUpdates(true);
      this.refreshConceptData(this.conceptId);
    } catch (err: any) {
      const msg = err?.error?.detail || err?.error?.message || err?.message || 'Failed to delete file';
      this.toastr.error(msg, 'Error');
    }
  }

  private generateAttachId(): string {
    return `att_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
  }

  private simulateUpload(cat: AttachCategory, entryId: string): void {
    let progress = 0;
    const interval = setInterval(() => {
      progress = Math.min(progress + 10, 100);

      // Match by stable id, not object reference — the array items are
      // replaced with new objects on every tick, so comparing against the
      // originally-captured object would only ever match once (tick 1)
      // and then silently stop updating, leaving progress stuck.
      this.attachments[cat] = this.attachments[cat].map(item =>
        item.id === entryId ? { ...item, progress } : item
      );
      this.cdr.detectChanges();

      if (progress >= 100) clearInterval(interval);
    }, 200);
  }

  // ── Supporting-doc upload ─────────────────────────────────────────────
  triggerFileUpload(): void { this.fileInput?.nativeElement.click(); }

  onFileSelected(event: Event): void {
    const input = event.target as HTMLInputElement;
    if (input.files) this.uploadedFiles = [...this.uploadedFiles, ...Array.from(input.files)];
  }

  onDragOver(event: DragEvent): void { event.preventDefault(); event.stopPropagation(); }

  onDrop(event: DragEvent): void {
    event.preventDefault(); event.stopPropagation();
    const files = event.dataTransfer?.files;
    if (files) this.uploadedFiles = [...this.uploadedFiles, ...Array.from(files)];
  }

  removeFile(index: number): void {
    this.uploadedFiles = this.uploadedFiles.filter((_, i) => i !== index);
  }

  formatFileSize(bytes: number): string {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1048576) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / 1048576).toFixed(1)} MB`;
  }

  // ── Breadcrumb ────────────────────────────────────────────────────────
  navigateToConcepts(): void { this.router.navigate(['/dashboard']); }

  // Concept IDs are assigned by the backend on creation — no client-side generation.

  // ── Client Approval ───────────────────────────────────────────────────
  async onApprovalSubmit(): Promise<void> {
    // Flips the Client Approval tab's own "submit was attempted" flag —
    // mirrors developmentSubmitAttempted in onSubmit(). Keeps a failed
    // Development submit from ever lighting up estimatedVolume/
    // estimatedDollars errors here, and vice versa.
    this.approvalSubmitAttempted = true;
    if (!this.canSubmitApproval) {
      this.toastr.error('You do not have permission to submit client approvals.', 'Access Denied');
      return;
    }
    if (this.isDraftConcept) {
      this.toastr.error(
        'This concept is still a draft. Please submit the Concept Development tab first before adding a Client Approval.',
        'Concept Not Submitted'
      );
      return;
    }
    if (!this.conceptId) {
      this.toastr.error(
        'Please save the Concept Information first before submitting Client Approval.',
        'Concept Not Saved'
      );
      return;
    }
    // Safety net behind the [disabled] binding on the button itself —
    // blocks the call even if it's triggered some other way (e.g. Enter
    // key) while a file's progress bar hasn't reached 100% yet.
    if (this.isApprovalAttachmentUploading) {
      this.toastr.error('Please wait for all files to finish uploading.', 'Upload in progress');
      return;
    }

    // Every field on this tab is mandatory before submitting.
    const requiredFields: { control: string; label: string }[] = [
      { control: 'clientConceptName', label: 'Concept Name' },
      { control: 'clientConceptDescription', label: 'Client Concept Description' },
      { control: 'clientApprovalStatus', label: 'Client Approval Status' },
      { control: 'submittedToClientOn', label: 'Submitted To Client On' },
      { control: 'clientEstimatedVolume', label: 'Estimated Volume' },
      { control: 'clientEstimatedDollars', label: 'Estimated Dollars' },
      { control: 'clientApprovalNotes', label: 'Client Review & Approval Notes' }
    ];

    const missing = requiredFields.filter(f => {
      const value = this.form.get(f.control)?.value;
      return this.isRequiredFieldMissing(f.control, value);
    });

    if (missing.length > 0) {
      // Mark touched so the template's *ngIf error messages light up too
      missing.forEach(f => this.form.get(f.control)?.markAsTouched());
      this.toastr.error(
        `Please fill in: ${missing.map(f => f.label).join(', ')}`,
        'Required fields missing'
      );
      return;
    }

    const submittedDateControl = this.form.get('submittedToClientOn');
    if (submittedDateControl?.hasError('invalidDateRange') || submittedDateControl?.hasError('pastDate')) {
      const isPast = submittedDateControl.hasError('pastDate');
      this.toastr.error(
        isPast
          ? 'Submitted To Client On cannot be in the past.'
          : 'Please enter a valid date (year must be between 1900 and 2100).',
        'Invalid Date'
      );
      submittedDateControl.markAsTouched();
      return;
    }

    // Attachment is optional on the Client Approval tab — skip the
    // "at least one file" guard that was here previously.

    // Safety net: never submit a blank client-facing concept name —
    // fall back to the main Concept Name if it was somehow left empty.
    this.ensureClientConceptName();

    // Nothing on this tab was actually touched since it was last saved —
    // updating would just write back exactly what's already saved. Block
    // it instead of hitting the backend for a no-op save with a false
    // "updated successfully" toast. Only applies once there's a prior save
    // to compare against — the very first submit always goes through.
    const isApprovalUpdate = this.clientApprovalCompleted === 1;
    if (isApprovalUpdate && !this.getApprovalFieldsChanged() && !this.getApprovalAttachmentsChanged()) {
      this.toastr.info('No changes to update.', 'Nothing to Save');
      return;
    }

    const approvalData = {
      conceptId: this.conceptId,
      conceptname: this.form.get('conceptName')?.value,
      clientConceptName: this.form.get('clientConceptName')?.value,
      clientConceptDescription: this.form.get('clientConceptDescription')?.value,
      clientApprovalStatus: this.form.get('clientApprovalStatus')?.value,
      submittedToClientOn: this.form.get('submittedToClientOn')?.value,
      clientApprovalNotes: this.form.get('clientApprovalNotes')?.value,
      estimatedVolume: this.form.get('clientEstimatedVolume')?.value,
      estimatedDollars: this.form.get('clientEstimatedDollars')?.value,
      clientApprovalCompleted: 1
    };
    console.log('Client Approval Data:', approvalData);


    await this.submitClientApproval(this.conceptId, approvalData);
  }

  async submitClientApproval(conceptId: string, data: any): Promise<void> {
    this.loading = true;
    const isUpdate = this.clientApprovalCompleted === 1;


    try {
      const user_id = Number(sessionStorage.getItem('userId'));
      const formData = new FormData();
      formData.append('data', JSON.stringify(data));
      formData.append('user_id', user_id.toString());
      formData.append('concept_id', conceptId || '');

      // Only newly-picked files have real bytes to upload — restored
      // placeholders (zero-byte, progress 100) are already saved on the
      // backend and are skipped here, same pattern as submitConcept()'s
      // filesToUpload above.
      const filesToUpload = this.attachments.approval.filter(
        f => f.file && f.file.size > 0
      );
      filesToUpload.forEach(entry => {
        formData.append('files', entry.file, entry.file.name);
        formData.append('categories', 'approval');
        formData.append('file_names', entry.file.name);
        formData.append('file_sizes', entry.file.size.toString());
      });

      // Diagnostic — remove once attachments are confirmed saving.
      console.log('[submitClientApproval] filesToUpload count:', filesToUpload.length,
        filesToUpload.map(f => f.name));

      const response = await this.service.submitclientApproval(formData, conceptId).toPromise();
      filesToUpload.forEach(entry => { entry.progress = 100; });
      this.clientApprovalCompleted = 1;
      // These fields are now exactly what's persisted — clear their dirty
      // flags so a subsequent Update click with no further edits is
      // correctly recognized as "no changes" (see getApprovalFieldsChanged()).
      this.approvalFields.forEach(f => this.form.get(f)?.markAsPristine());

      this.toastr.success(
        isUpdate
          ? 'Approval updated successfully!'
          : 'Approval submitted successfully!',
        'Success'
      );

      // Refresh the left panel + soft-reload this concept's data so the
      // page reflects exactly what was just persisted — no full page
      // reload. Quiet background re-sync (see loadLatestUpdates()) so the
      // list doesn't blink. Deliberately leaves resnapshotDevelopmentFields
      // at its default (false): this endpoint can also persist the shared
      // estimatedVolume/estimatedDollars fields, but that must NOT be
      // allowed to clear an unresolved "add a Development Note" state on
      // the Concept Development tab — see fetchAndApplyConcept().
      this.loadLatestUpdates(true);
      this.refreshConceptData(conceptId);
    } catch (error) {
      // this.toastr.error('Error submitting approval.', 'Error');
      this.toastr.error(
        'Please save the concept first before  submitting  client Approval.',
        'Concept not saved'
      );
    } finally {
      this.loading = false;
    }
  }
  // ── Supporting Documents ──────────────────────────────────────────────
  onAddSupportingDoc(): void {
    if (this.blockIfCannotManageDocs()) return;
    this.supportingDocs = [
      ...this.supportingDocs,
      this.blankDoc()
    ];
  }

  /** Same split as removeAttachment() above: a slot added/filled this
   *  session but never submitted has no backend record (no
   *  attachmentId) — just drop it from the array. A restored slot always
   *  has one (every active_files row, even URL-only docs, gets an
   *  AttachmentId — see patchSupportingDocs), so it has to be deleted on
   *  the backend first, and the card only disappears once that succeeds.
   *
   *  NOTE: this removes the SLOT (and retires its docIndex — that
   *  docIndex is never reused, see allocateDocIndex()). It intentionally
   *  does NOT renumber/reassign docIndex on the remaining docs — their
   *  identity must stay exactly what it was before this deletion. */
  async removeSupportingDoc(index: number): Promise<void> {
    if (this.blockIfCannotManageDocs()) return;
    const doc = this.supportingDocs[index];

    if (!doc?.attachmentId) {
      this.removeSupportingDocLocally(index);
      return;
    }

    // NEW: once a concept has supporting documents on the backend, at
    // least one must remain — block removing the last saved one instead
    // of letting the delete API fire immediately and leave the concept
    // with zero supporting documents.
    const savedDocsCount = this.supportingDocs.filter(d => d.attachmentId).length;
    if (savedDocsCount <= 1) {
      this.toastr.error(
        'At least one supporting document is required. Add a replacement before removing this one.',
        'Cannot Remove'
      );
      return;
    }

    const user_id = Number(sessionStorage.getItem('userId'));
    try {
      const formData = new FormData();
      formData.append('concept_id', this.conceptId);
      formData.append('attachment_id', doc.attachmentId.toString());
      formData.append('category', 'supporting_docs');
      formData.append('user_id', user_id.toString());

      await this.service.deleteattachment(formData).toPromise();
      this.removeSupportingDocLocally(index);
      this.toastr.success('Document deleted successfully!', 'Success');
      this.loadLatestUpdates(true);
      this.refreshConceptData(this.conceptId);
    } catch (err: any) {
      const msg = err?.error?.detail || err?.error?.message || err?.message || 'Failed to delete document';
      this.toastr.error(msg, 'Error');
    }
  }

  private removeSupportingDocLocally(index: number): void {
    this.supportingDocs = this.supportingDocs.filter((_, i) => i !== index);
    if (this.modalSelectedIndex !== null) {
      if (index === this.modalSelectedIndex) this.modalSelectedIndex = null;
      else if (index < this.modalSelectedIndex) this.modalSelectedIndex--;
    }
  }

  triggerDocFileInput(index: number): void {
    if (this.blockIfCannotManageDocs()) return;
    this.pendingDocIndex = index;
    this.docFileInputRef?.nativeElement.click();
  }

  onGlobalDocFileSelected(event: Event): void {
    if (this.blockIfCannotManageDocs()) return;
    if (this.pendingDocIndex === null) return;

    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];

    if (!file) return;
    // Validate file type
    if (!this.isValidFile(file)) {
      this.toastr.error(
        'Only PDF, Word (.doc/.docx), and Excel (.xls/.xlsx) files are allowed.',
        'Invalid File Type'
      );
      input.value = '';
      this.pendingDocIndex = null;
      return;
    }

    const idx = this.pendingDocIndex;

    // Check for duplicate filenames in other Supporting Document slots
    const isDuplicate = this.supportingDocs.some(
      (doc, i) =>
        i !== idx &&
        doc.file &&
        doc.file.name.trim().toLowerCase() === file.name.trim().toLowerCase()
    );

    if (isDuplicate) {
      this.toastr.error(
        `"${file.name}" is already attached in another Supporting Document slot.`,
        'Duplicate File'
      );
      input.value = '';
      this.pendingDocIndex = null;
      return;
    }

    // Keep sourceurl unchanged and update the selected document
    this.supportingDocs = this.supportingDocs.map((doc, i) =>
      i === idx
        ? {
          ...doc,
          name: file.name,
          file,
          uploadProgress: 0,
          originalFileName: undefined,
          originalFileSize: undefined
        }
        : doc
    );

    this.cdr.detectChanges();
    this.simulateDocUpload(idx);

    input.value = '';
    this.pendingDocIndex = null;
  }

  onDocFileSelected(event: Event, index: number): void {
    if (this.blockIfCannotManageDocs()) return;
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;

    const isDuplicate = this.supportingDocs.some(
      (doc, i) => i !== index && doc.file && doc.file.name.trim().toLowerCase() === file.name.trim().toLowerCase()
    );
    if (isDuplicate) {
      this.toastr.error(
        `"${file.name}" is already attached in another Supporting Document slot.`,
        'Duplicate File'
      );
      input.value = '';
      return;
    }

    // sourceurl left as-is — see the matching note in onGlobalDocFileSelected.
    const updated = [...this.supportingDocs];
    updated[index] = { ...updated[index], name: file.name.replace(/\.[^.]+$/, ''), file, uploadProgress: 0, originalFileName: undefined, originalFileSize: undefined };
    this.supportingDocs = updated;
    this.simulateDocUpload(index);
    input.value = '';
  }

  onDocDrop(event: DragEvent, index: number): void {
    event.preventDefault(); event.stopPropagation();
    if (this.blockIfCannotManageDocs()) return;
    const file = event.dataTransfer?.files?.[0];
    if (!file) return;

    const isDuplicate = this.supportingDocs.some(
      (doc, i) => i !== index && doc.file && doc.file.name.trim().toLowerCase() === file.name.trim().toLowerCase()
    );
    if (isDuplicate) {
      this.toastr.error(
        `"${file.name}" is already attached in another Supporting Document slot.`,
        'Duplicate File'
      );
      return;
    }

    // sourceurl left as-is — see the matching note in onGlobalDocFileSelected.
    const updated = [...this.supportingDocs];
    updated[index] = { ...updated[index], name: file.name.replace(/\.[^.]+$/, ''), file, uploadProgress: 0, originalFileName: undefined, originalFileSize: undefined };
    this.supportingDocs = updated;
    this.simulateDocUpload(index);
  }

  /** Clears this entire Supporting Document card back to blank — file,
   *  Document Name, Source URL, and PDF Location all reset, same as an
   *  empty slot. The card itself (and its position in the list) stays;
   *  use the "Delete" button (removeSupportingDoc()) to remove the slot
   *  entirely.
   *
   *  This is UI-only: it does NOT call the delete API, so if this slot
   *  was restored from a saved concept, the backend record isn't
   *  touched yet. originalFileName/originalFileSize are cleared too —
   *  otherwise resolveDocFileMeta() would fall back to the old file info
   *  on the next submit and silently resurrect a file the user just
   *  cleared from the screen. attachmentId is kept so the "Delete"
   *  button can still remove the underlying backend record if the user
   *  uses it afterward. docIndex is ALSO kept — this slot's identity
   *  doesn't change just because its contents were cleared. */
  clearDocFile(event: Event, index: number): void {
    event.stopPropagation();
    if (this.blockIfCannotManageDocs()) return;
    const updated = [...this.supportingDocs];
    updated[index] = {
      ...updated[index],
      name: '',
      sourceurl: '',
      pdfLocation: '',
      file: null,
      uploadProgress: 0,
      downloadUrl: undefined,
      restored: false,
      originalFileName: undefined,
      originalFileSize: undefined
    };
    this.supportingDocs = updated;
  }

  // ── Doc Viewer ────────────────────────────────────────────────────────
  isWordExt(ext: string): boolean { return ['docx', 'doc'].includes(ext.toLowerCase()); }
  isExcelExt(ext: string): boolean { return ['xlsx', 'xls'].includes(ext.toLowerCase()); }
  isPdfExt(ext: string): boolean { return ext.toLowerCase() === 'pdf'; }

  /** Fetches a stored file's bytes from /api/download-attachment/{id} and
   *  opens it in the in-app preview (Word/Excel/PDF viewer). Shared by
   *  both the Concept Development "Attachments" cards (viewAttachment)
   *  and Supporting Documents (viewDoc), since both now resolve to the
   *  same backend endpoint, keyed by AttachmentId. */
  private fetchAndViewRemoteFile(name: string, downloadUrl: string): void {
    // this.toastr.info('Loading file…', '', { timeOut: 1500 });

    this.service.getPdfFile(downloadUrl).subscribe({
      next: (blob: Blob) => {
        // new File() does NOT inherit the blob's MIME type unless told to.
        // Without it, the blob: URL built later in renderDocPdf() ends up
        // with no type at all, so the browser can't tell it's a PDF and
        // falls back to showing the raw bytes as plain text instead of
        // rendering it — falling back to the extension when the server
        // didn't set (or HttpClient didn't preserve) a Content-Type.
        const mimeType = blob.type || this.inferMimeType(name);
        const realFile = new File([blob], name, { type: mimeType });
        this.viewDoc({ name, sourceurl: '', pdfLocation: '', uploadProgress: 100, file: realFile, docIndex: -1 });
      },
      error: (err) => {
        console.error('Failed to fetch file:', err);
        this.toastr.error('Could not load file for preview', 'Error');
      }
    });
  }

  private inferMimeType(name: string): string {
    const ext = name.split('.').pop()?.toLowerCase() ?? '';
    switch (ext) {
      case 'pdf': return 'application/pdf';
      case 'doc': return 'application/msword';
      case 'docx': return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
      case 'xls': return 'application/vnd.ms-excel';
      case 'xlsx': return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
      default: return 'application/octet-stream';
    }
  }

  /** View an Attachments-tab file (specs/table/other) using the same
   *  Word/Excel preview popup as Supporting Documents. Newly-selected
   *  files already have real bytes in memory; files restored from a
   *  saved concept are empty placeholders and need to be fetched first. */
  async viewAttachment(f: AttachFile): Promise<void> {
    if (f.progress < 100) {
      this.toastr.error('Wait for the upload to finish before viewing', 'Error');
      return;
    }

    // Already have real bytes (just-selected/uploaded file) — view directly.
    if (f.file && f.file.size > 0) {
      this.viewDoc({ name: f.name, sourceurl: '', pdfLocation: '', uploadProgress: f.progress, file: f.file, docIndex: -1 });
      return;
    }

    // Placeholder from a saved concept — fetch the actual bytes first.
    if (!f.downloadUrl) {
      this.toastr.error('No file available to preview', 'Error');
      return;
    }

    this.fetchAndViewRemoteFile(f.name, f.downloadUrl);
  }

  async viewDoc(doc: SupportingDoc): Promise<void> {
    if (doc.file && doc.uploadProgress < 100) {
      this.toastr.error('Wait for the upload to finish before viewing', 'Error');
      return;
    }

    // Restored from a saved concept — placeholder has no real bytes yet
    // (patchSupportingDocs sets a zero-byte File), fetch them from
    // /api/download-attachment/{id} first, same as viewAttachment.
    if ((!doc.file || doc.file.size === 0) && doc.downloadUrl) {
      this.fetchAndViewRemoteFile(doc.name, doc.downloadUrl);
      return;
    }

    const file = doc.file;
    if (file) {
      const ext = file.name.split('.').pop()?.toLowerCase() ?? '';

      if (this.isWordExt(ext) || this.isExcelExt(ext) || this.isPdfExt(ext)) {
        this.docViewerFileName = file.name;
        this.docViewerFileExt = ext.toUpperCase();
        this.docWordHtml = '';
        this.docSheets = [];
        this.docActiveSheet = 0;
        this.revokePdfUrl();
        this.docViewerVisible = true;
        this.docViewerLoading = true;
        // Force a render now so the backdrop + "Loading preview…" state is
        // visible immediately, before we hand off to native browser APIs
        // below (createObjectURL / Blob.arrayBuffer) whose promise
        // resolution isn't always picked up by zone.js, which is what was
        // causing the popup to look "stuck" until a second click.
        this.cdr.detectChanges();
        try {
          if (this.isPdfExt(ext)) {
            this.renderDocPdf(file);
          } else {
            const arrayBuffer = await file.arrayBuffer();
            if (this.isWordExt(ext)) await this.renderDocWord(arrayBuffer);
            else this.renderDocExcel(arrayBuffer);
          }
        } catch (err: any) {
          this.toastr.error('Could not open file: ' + err.message, 'Error');
          this.docViewerVisible = false;
        } finally {
          this.docViewerLoading = false;
          // Native APIs above (Blob.arrayBuffer, createObjectURL) can
          // resolve outside Angular's zone, so explicitly trigger a render
          // here instead of waiting for the next zone-tracked event.
          this.cdr.detectChanges();
        }
        return;
      }
    }
    if (doc.pdfLocation) window.open(doc.pdfLocation, '_blank');
    else if (doc.sourceurl) window.open(doc.sourceurl, '_blank');
    else if (file && file.size > 0) {
      // Generic fallback for file types we don't render inline (images, etc.)
      const objectUrl = URL.createObjectURL(file);
      window.open(objectUrl, '_blank');
    }
    else this.toastr.error('Nothing to preview for this document', 'Error');
  }

  private renderDocPdf(file: File): void {
    this.docPdfObjectUrl = URL.createObjectURL(file);
    this.docPdfUrl = this.sanitizer.bypassSecurityTrustResourceUrl(this.docPdfObjectUrl);
  }

  private revokePdfUrl(): void {
    if (this.docPdfObjectUrl) {
      URL.revokeObjectURL(this.docPdfObjectUrl);
      this.docPdfObjectUrl = null;
    }
    this.docPdfUrl = null;
  }

  private async renderDocWord(buffer: ArrayBuffer): Promise<void> {
    const result = await mammoth.convertToHtml({ arrayBuffer: buffer });
    this.docWordHtml = this.sanitizer.bypassSecurityTrustHtml(result.value);
  }

  private renderDocExcel(buffer: ArrayBuffer): void {
    const workbook = XLSX.read(new Uint8Array(buffer), { type: 'array' });
    this.docSheets = workbook.SheetNames.map(name => {
      const ws = workbook.Sheets[name];
      const data: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });
      return { sheetName: name, headers: (data[0] as string[]) || [], rows: data.slice(1) };
    });
  }

  closeDocViewer(): void {
    this.docViewerVisible = false;
    this.revokePdfUrl();
  }
  setDocActiveSheet(i: number): void { this.docActiveSheet = i; }

  // ── Downloads ────────────────────────────────────────────────────────
  /** Forces the browser to save `blobOrFile` as `filename` instead of
   *  navigating to / opening it inline. Shared by every per-file Download
   *  button below and by both "Download All" actions. */
  private triggerDownload(blobOrFile: Blob, filename: string): void {
    const objectUrl = URL.createObjectURL(blobOrFile);
    const link = document.createElement('a');
    link.href = objectUrl;
    link.download = filename || 'download';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    // Give the browser a moment to pick up the blob: URL before revoking it.
    setTimeout(() => URL.revokeObjectURL(objectUrl), 1000);
  }

  /** Downloads a single Attachments-tab file (specs/table/other). Files
   *  just picked in this session already have real bytes in memory;
   *  files restored from a saved concept are zero-byte placeholders and
   *  need their bytes fetched from /api/download-attachment/{id} first —
   *  same split used by viewAttachment() above. */
  downloadAttachment(f: AttachFile): void {
    if (f.progress < 100) {
      this.toastr.error('Wait for the upload to finish before downloading', 'Error');
      return;
    }
    if (f.file && f.file.size > 0) {
      this.triggerDownload(f.file, f.name);
      return;
    }
    if (!f.downloadUrl) {
      this.toastr.error('No file available to download', 'Error');
      return;
    }
    this.service.getPdfFile(f.downloadUrl).subscribe({
      next: (blob: Blob) => this.triggerDownload(blob, f.name),
      error: (err) => {
        console.error('Failed to download file:', err);
        this.toastr.error('Could not download file', 'Error');
      }
    });
  }

  /** True once at least one file exists in any of the three Attachments
   *  categories — drives [disabled] on the card's "Download All" button. */
  get hasAnyAttachments(): boolean {
    return this.attachments.specs.length > 0 ||
      this.attachments.table.length > 0 ||
      this.attachments.other.length > 0;
  }

  /** Same as hasAnyAttachments, scoped to the Client Approval tab's own
   *  attachment category — drives [disabled] on that section's own
   *  "Download All" button without mixing in Development-tab files. */
  get hasAnyApprovalAttachments(): boolean {
    return this.attachments.approval.length > 0;
  }

  /** SPECS is the only required attachment category on the Concept
   *  Development tab — drives [disabled] on "Create Concept" / "Update"
   *  so the button stays disabled until at least one SPECS file is
   *  attached. */
  get hasSpecsAttachment(): boolean {
    return this.attachments.specs.length > 0;
  }

  /** Downloads every Attachments-tab file across all three categories.
   *  Browsers throttle/block several programmatic downloads fired in the
   *  same tick, so each one is staggered slightly. */
  downloadAllAttachments(): void {
    const all: AttachFile[] = [
      ...this.attachments.specs,
      ...this.attachments.table,
      ...this.attachments.other
    ];
    const ready = all.filter(f => f.progress >= 100 && (f.file?.size || f.downloadUrl));
    if (ready.length === 0) {
      this.toastr.error('No files available to download', 'Error');
      return;
    }
    ready.forEach((f, i) => setTimeout(() => this.downloadAttachment(f), i * 400));
  }

  /** Downloads every file attached in the Client Approval tab's own
   *  Attachments section. Mirrors downloadAllAttachments() above but
   *  scoped to the 'approval' category only. */
  downloadAllApprovalAttachments(): void {
    const ready = this.attachments.approval.filter(
      f => f.progress >= 100 && (f.file?.size || f.downloadUrl)
    );
    if (ready.length === 0) {
      this.toastr.error('No files available to download', 'Error');
      return;
    }
    ready.forEach((f, i) => setTimeout(() => this.downloadAttachment(f), i * 400));
  }

  /** Downloads a single Supporting Documents card's actual attached file.
   *  Source URL / PDF Location are reference metadata only — they are
   *  never opened or force-downloaded from here. A card with no real
   *  file behind it (no local File, no downloadUrl) has nothing to
   *  download, so this shows an error instead of navigating anywhere. */
  downloadDoc(doc: SupportingDoc): void {
    if (doc.file && doc.uploadProgress < 100) {
      this.toastr.error('Wait for the upload to finish before downloading', 'Error');
      return;
    }
    if (doc.file && doc.file.size > 0) {
      this.triggerDownload(doc.file, doc.name || doc.file.name);
      return;
    }
    if (doc.downloadUrl) {
      this.service.getPdfFile(doc.downloadUrl).subscribe({
        next: (blob: Blob) => this.triggerDownload(blob, doc.name || 'document'),
        error: (err) => {
          console.error('Failed to download file:', err);
          this.toastr.error('Could not download file', 'Error');
        }
      });
      return;
    }
    this.toastr.error('No file attached to download', 'Error');
  }

  /** True once at least one Supporting Document card has an actual
   *  downloadable file behind it — drives [disabled] on the section's
   *  "Download All" button. Deliberately does NOT count sourceurl/
   *  pdfLocation alone: those are reference links, not files, and
   *  downloadDoc() itself refuses to download a card that only has one
   *  of those — so counting them here used to leave "Download All"
   *  enabled even when there was nothing it could actually download,
   *  producing a per-row error the moment it ran. */
  get hasAnySupportingDocs(): boolean {
    return this.supportingDocs.some(d => d.file || d.downloadUrl);
  }

  /** Downloads every Supporting Document card that has an actual file
   *  attached. Staggered for the same reason as downloadAllAttachments()
   *  above. Cards with only a Source URL or PDF Location (no real file)
   *  are skipped — see hasAnySupportingDocs above for why. */
  downloadAllSupportingDocs(): void {
    const ready = this.supportingDocs.filter(d => d.file || d.downloadUrl);
    if (ready.length === 0) {
      this.toastr.error('No documents available to download', 'Error');
      return;
    }
    ready.forEach((d, i) => setTimeout(() => this.downloadDoc(d), i * 400));
  }

  private simulateDocUpload(index: number): void {
    // Resolve by the slot's stable docIndex, not by its array position at
    // call time. Array position shifts whenever any doc is deleted
    // (Array.filter in removeSupportingDocLocally) — if this interval kept
    // writing to a raw numeric index across that shift, it would start
    // overwriting a DIFFERENT card's data every tick and force-reassign
    // `supportingDocs` to a new array reference each time, which makes the
    // whole *ngFor list re-render/flicker. Tracking by docIndex means a
    // deletion elsewhere simply has no effect on this upload.
    const targetDocIndex = this.supportingDocs[index]?.docIndex;
    if (targetDocIndex === undefined) return;

    const findCurrentIndex = (): number =>
      this.supportingDocs.findIndex(d => d.docIndex === targetDocIndex);

    let progress = 0;
    const interval = setInterval(() => {
      const i = findCurrentIndex();
      if (i === -1) {
        // The slot this upload belongs to was deleted mid-upload — stop
        // silently instead of writing into whatever now occupies the old
        // array position.
        clearInterval(interval);
        return;
      }

      progress += 5;
      this.supportingDocs[i] = { ...this.supportingDocs[i], uploadProgress: progress };
      this.supportingDocs = [...this.supportingDocs];
      this.cdr.detectChanges();
      if (progress >= 100) {
        clearInterval(interval);

        // Show the success banner immediately on completion, then auto-hide
        // it 5s later — the rest of the card (icon, name, source URL, etc.)
        // stays exactly as-is, only the banner disappears.
        const doneIndex = findCurrentIndex();
        if (doneIndex !== -1 && this.supportingDocs[doneIndex]) {
          this.supportingDocs[doneIndex] = { ...this.supportingDocs[doneIndex], successBannerVisible: true };
          this.supportingDocs = [...this.supportingDocs];
          this.cdr.detectChanges();
        }
        setTimeout(() => {
          const bannerIndex = findCurrentIndex();
          if (bannerIndex !== -1 && this.supportingDocs[bannerIndex]) {
            this.supportingDocs[bannerIndex] = { ...this.supportingDocs[bannerIndex], successBannerVisible: false };
            this.supportingDocs = [...this.supportingDocs];
            this.cdr.detectChanges();
          }
        }, 5000);
      }
    }, 300);
  }

  // ── Upload Modal ──────────────────────────────────────────────────────
  openUploadModal(): void {
    this.showUploadModal = true;
    this.modalSelectedIndex = this.supportingDocs.length > 0 ? this.supportingDocs.length - 1 : null;
  }

  closeUploadModal(): void { this.showUploadModal = false; }

  onModalFileSelected(event: Event): void {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;
    const newDoc: SupportingDoc = { ...this.blankDoc(), name: file.name.replace(/\.[^.]+$/, ''), file };
    this.supportingDocs = [...this.supportingDocs, newDoc];
    this.modalSelectedIndex = this.supportingDocs.length - 1;
    this.simulateDocUpload(this.modalSelectedIndex);
    input.value = '';
  }

  onModalDrop(event: DragEvent): void {
    event.preventDefault(); event.stopPropagation();
    const file = event.dataTransfer?.files?.[0];
    if (!file) return;
    const newDoc: SupportingDoc = { ...this.blankDoc(), name: file.name.replace(/\.[^.]+$/, ''), file };
    this.supportingDocs = [...this.supportingDocs, newDoc];
    this.modalSelectedIndex = this.supportingDocs.length - 1;
    this.simulateDocUpload(this.modalSelectedIndex);
  }

  /** Resolves the fileName/fileSize to report to the backend for one
   *  Supporting Document slot's update:
   *  - A freshly-picked file (real bytes, chosen this session) reports its
   *    own name/size — the binary itself goes out in the same request.
   *  - A restored slot the user didn't touch has no binary to attach
   *    (its `file` is just a zero-byte placeholder — see
   *    patchSupportingDocs), so it reports back the fileName/fileSize it
   *    was originally restored with, so the backend knows this slot's
   *    existing file is still intact rather than reading "no file info"
   *    as "the file was removed."
   *  - Anything else (URL-only / empty slot) has no file info to give. */
  private resolveDocFileMeta(doc: SupportingDoc): { fileName?: string; fileSize?: number } {
    if (doc.file && doc.file.size > 0) {
      return { fileName: doc.file.name, fileSize: doc.file.size };
    }
    if (doc.originalFileName) {
      return { fileName: doc.originalFileName, fileSize: doc.originalFileSize };
    }
    return {};
  }

  async onDocSubmit(): Promise<void> {
    const isUpdate = this.supportingDocumentsCompleted === 1;

    if (!this.canManageSupportingDocs) {
      this.toastr.error('You do not have permission to submit supporting documents.', 'Access Denied');
      return;
    }
    if (this.isDraftConcept) {
      this.toastr.error(
        'This concept is still a draft. Please submit the Concept Development tab first before adding Supporting Documents.',
        'Concept Not Submitted'
      );
      return;
    }
    if (this.isSupportingDocUploading) {
      this.toastr.error('Please wait for all files to finish uploading.', 'Upload in progress');
      return;
    }
    const user_id = Number(sessionStorage.getItem('userId'));
    if (!this.conceptId) {
      this.toastr.error(
        'Please save the concept first before submitting supporting documents.',
        'Concept not saved'
      );
      return;
    }

    // Only include docs that have a file, URL, or PDF location filled in.
    const validDocs = this.supportingDocs.filter(
      d => d.file || d.sourceurl?.trim() || d.pdfLocation?.trim()
    );

    if (validDocs.length === 0) {
      this.toastr.error(
        isUpdate
          ? 'Please add at least one supporting document before updating.'
          : 'Please add at least one supporting document before submitting.',
        'No documents'
      );
      return;
    }

    // A Source URL / PDF Location is metadata ABOUT an attached file, not
    // a standalone substitute for one — a doc slot can't carry a link with
    // no file behind it (they're saved together as one set). This mainly
    // catches: clear a slot's existing file+link, then type in only a new
    // link and hit Update. Without this check the slot still "looks valid"
    // (sourceurl is non-empty) and gets submitted with no fileName/fileSize
    // — which reads to the backend as "this slot's existing file is still
    // intact" (see resolveDocFileMeta()) rather than "the file was removed
    // and only a link is left," so the old file/green indicator silently
    // persists even though the user meant to replace it with just a link.
    const linkOnlyDocs = validDocs.filter(d => {
      const hasLink = !!(d.sourceurl?.trim() || d.pdfLocation?.trim());
      const hasFile = !!(d.file && d.file.size > 0) || !!d.originalFileName;
      return hasLink && !hasFile;
    });
    if (linkOnlyDocs.length > 0) {
      this.toastr.error(
        'A Source URL or PDF Location can only be saved together with an attached file. Please attach a file before adding a link.',
        'File Required'
      );
      return;
    }

    // Nothing on this tab was actually touched since it was last saved —
    // updating would just write back exactly what's already saved. Block
    // it instead of hitting the backend for a no-op save with a false
    // "updated successfully" toast. Only applies once there's a prior save
    // to compare against — the very first submit always goes through.
    if (isUpdate && !this.getSupportingDocsChanged()) {
      this.toastr.info('No changes to update.', 'Nothing to Save');
      return;
    }

    this.loading = true;

    try {
      // Lean metadata — only Supporting Documents payload + completion flags.
      // fileName/fileSize is resolved per-doc below: a freshly-picked file
      // reports its own name/size (and its binary goes out alongside, see
      // the forEach below); a restored-but-untouched slot has no binary to
      // attach, so it reports back the fileName/fileSize it was restored
      // with — without this, an unmodified slot looks file-less to the
      // backend on every resubmit and its existing file association gets
      // wiped, even though nothing about it actually changed.
      //
      // IMPORTANT: docIndex is each slot's STABLE identity (see the
      // SupportingDoc.docIndex doc comment) and is sent explicitly here.
      // It must NEVER be derived from this array's position (e.g. via
      // validDocs.map((d, i) => ...)) — array position shifts every time
      // a doc is deleted, which previously caused a later doc to be sent
      // under an earlier (deleted) doc's index and silently deactivate
      // the wrong attachment on the backend.
      const sdMetadata = {
        concept_id: this.conceptId,
        SupportingDocumentsCompleted: 1,
        supportingDocs: validDocs.map(d => {
          const { fileName, fileSize } = this.resolveDocFileMeta(d);
          return {
            docIndex: d.docIndex,
            name: d.name,
            sourceurl: d.sourceurl,
            pdfLocation: d.pdfLocation,
            fileName,
            fileSize
          };
        })
      };

      // Single combined request: metadata + every doc's binary file (for
      // whichever docs actually have one) go out together in ONE
      // FormData/API call, instead of one call per doc (or per doc per
      // chunk) like before. 'doc_indices' ties each entry in 'files' back
      // to its STABLE docIndex (see above) — NOT its position in
      // validDocs/this loop.
      const formData = new FormData();
      formData.append('concept_id', this.conceptId);
      formData.append('metadata', JSON.stringify(sdMetadata));
      formData.append('category', 'supporting_docs');
      formData.append('user_id', user_id.toString());

      validDocs.forEach((doc) => {
        if (!doc.file || doc.file.size === 0) return; // URL-only doc — nothing to upload
        formData.append('files', doc.file, doc.file.name);
        formData.append('doc_indices', doc.docIndex.toString());
        formData.append('doc_names', doc.name || doc.file.name);
        formData.append('source_urls', doc.sourceurl || '');
        formData.append('file_names', doc.file.name);
        formData.append('file_sizes', doc.file.size.toString());
      });

      const res: any = await this.service.submitsupportingdocuments(formData).toPromise();
      if (res?.status === 'error' || res?.detail) {
        this.toastr.error(res?.detail || res?.message || 'Something went wrong', 'Error');
        return;
      }

      // Mark every doc complete now that the single request has succeeded
      // (no per-chunk progress to track anymore).
      validDocs.forEach(doc => {
        const globalIndex = this.supportingDocs.indexOf(doc);
        if (globalIndex !== -1) {
          this.supportingDocs[globalIndex] = { ...this.supportingDocs[globalIndex], uploadProgress: 100 };
        }
      });
      this.supportingDocs = [...this.supportingDocs];

      this.supportingDocumentsCompleted = 1;
      this.toastr.success(
        isUpdate
          ? 'Supporting documents updated successfully!'
          : 'Supporting documents submitted successfully!',
        'Success'
      );

      // Refresh the left panel + soft-reload this concept's data so the
      // page reflects exactly what was just persisted — no full page
      // reload. Quiet background re-sync (see loadLatestUpdates()) so the
      // list doesn't blink.
      this.loadLatestUpdates(true);
      this.refreshConceptData(this.conceptId);

    } catch (err: any) {
      const msg =
        err?.error?.detail ||
        err?.error?.message ||
        err?.message ||
        'Upload failed. Please try again.';
      this.toastr.error(this.friendlyErrorMessage(msg), 'Error');
    } finally {
      this.loading = false;
    }
  }

  // ── Latest Updates list ────────────────────────────────────────────────
  // Shows EVERY concept in the system, regardless of who created it or
  // what role is currently logged in — this is the shared, read-only
  // visibility layer everyone gets, separate from the edit restrictions
  // enforced elsewhere (applyRoleRestrictions / refreshAllowedStatuses).
  // Uses /api/latest-updates (unfiltered, Concepts UNION active
  // ConceptDrafts) rather than /api/dashboard-concepts, which only reads
  // the Concepts table and silently omits anything still saved as a
  // draft — or getConceptsByUserId, which only returns the current
  // user's own concepts and drafts.
  /** @param background when true, this is a quiet re-sync after some other
   *  save/submit succeeded (Save as Draft, Submit, Client Approval,
   *  Supporting Documents) — the panel already has the list on screen, so
   *  we must NOT flip latestUpdatesLoading on/off around the call. Doing
   *  so used to swap the entire *ngIf="!latestUpdatesLoading" list out for
   *  the spinner and back again for every single save, which is exactly
   *  what read as the Concept List "blinking" whenever the user clicked
   *  Save as Draft (or Submit, or either of the other tabs' Submits).
   *  Only the true first-ever load (ngOnInit) needs the spinner. */
  private loadLatestUpdates(background: boolean = false): void {
    if (!background) {
      this.latestUpdatesLoading = true;
    }

    this.service.getLatestUpdates().subscribe({
      next: (res) => {
        const concepts: any[] = res?.data ?? [];

        this.latestConcepts = concepts
          .slice()
          .sort((a, b) => {
            const bTime = new Date(b.UpdatedDate ?? b.CreatedDate).getTime();
            const aTime = new Date(a.UpdatedDate ?? a.CreatedDate).getTime();
            return bTime - aTime;
          })
          .map(c => ({
            ...c,
            statusClass: this.getStatusClass(c.DevelopmentStatus),
            isDraft: !!(c.IsDraft ?? c.isDraft)
          }));

        if (!background) {
          this.latestUpdatesLoading = false;
        }

        // Bring the concept the user is currently working on into view —
        // its position in this list can shift (e.g. after an edit bumps it
        // via UpdatedDate), so without this it can silently scroll out of
        // the visible panel even though .active styling is still correctly
        // applied to its card.
        this.scrollActiveConceptIntoView();
      },
      error: (err) => {
        console.error('Failed to load latest updates:', err);
        if (!background) {
          this.latestUpdatesLoading = false;
        }
      }
    });
  }

  /** Scrolls the currently-open concept's card into view (aligned to the
   *  top of the panel) in the Latest Updates panel, if it isn't already
   *  visible. Runs after every list (re)load — initial load, background
   *  refresh after a save, and concept switches — since the active card's
   *  position can change (new sort order, list re-fetch) independently of
   *  the user scrolling anywhere.
   *
   *  On a brand-new navigation into this page (e.g. clicking a row on the
   *  Dashboard), this can be called BEFORE the Latest Updates list has
   *  finished loading — loadConcept()'s single-concept fetch and
   *  loadLatestUpdates()'s full-list fetch race, and whichever call lands
   *  first won't find the card in the DOM yet. Rather than rely solely on
   *  the other call's own scrollActiveConceptIntoView() to pick up the
   *  slack, this retries for a couple of seconds until the card actually
   *  exists, so the highlight+scroll always lands correctly regardless of
   *  which fetch resolves first. */
  private scrollActiveConceptIntoView(): void {
    if (!this.conceptId) return;
    const targetId = this.conceptId;
    let attempts = 0;
    const tryScroll = () => {
      // Bail if the user has since navigated to a different concept.
      if (this.conceptId !== targetId) return;
      const el = document.querySelector(
        `[data-concept-id="${targetId}"], [data-current-concept-id="${targetId}"]`
      );
      if (el) {
        el.scrollIntoView({ block: 'start', behavior: 'smooth' });
        return;
      }
      attempts++;
      if (attempts < 20) {
        setTimeout(tryScroll, 100);
      }
    };
    // Wait a tick for the *ngFor to actually render before the first check.
    setTimeout(tryScroll);
  }

  /** Resets the scrollable form area (.tab-content) back to the top whenever
   *  a concept is (re)loaded. Without this, navigating to a concept — e.g.
   *  clicking a row on the Dashboard, or picking a different concept from
   *  the Latest Updates panel — leaves the form wherever the PREVIOUS
   *  concept happened to be scrolled to, since Angular reuses this same
   *  component instance across /concept-create/:id navigations (see
   *  routeSub in ngOnInit) and never remounts .tab-content. */
  private scrollTabContentToTop(): void {
    setTimeout(() => {
      const el = this.tabContentRef?.nativeElement;
      if (el) el.scrollTop = 0;
    });
  }

  /** trackBy for the Concept List *ngFor — keyed on the stable anchor id,
   *  so a background refresh (see loadLatestUpdates()'s background param)
   *  only patches the rows that actually changed instead of Angular
   *  tearing down and rebuilding every card in the list each time. */
  trackByConceptId(index: number, item: LatestConceptItem): string {
    return item.ConceptId;
  }


  getStatusClass(status: string): string {
    switch (status) {
      case 'New':
        return 'lu-new';

      case 'Programming Queue':
      case 'Programming':
      case 'Researching':
        return 'lu-progress';

      case 'Result Set QA':
      case 'QA Revise':
        return 'lu-qa';

      case 'Approved':
        return 'lu-approved';

      case 'Client Review':
      case 'Client Revise':
      case 'Client Resubmit':
        return 'lu-client-review';

      case 'Client Approved':
        return 'lu-client-approved';

      case 'Client Denied':
        return 'lu-client-denied';

      case 'Pre-Production':
        return 'lu-preproduction';

      case 'Production':
        return 'lu-production';

      case 'Closed':
        return 'lu-closed';

      case 'Hold':
      case 'Revisit':
        return 'lu-hold';

      case 'Superseded':
        return 'lu-superseded';

      default:
        return '';
    }
  }
  // onSelectLatestConcept(item: LatestConceptItem): void {
  //   if (item.conceptId === this.conceptId) return;
  //   this.isEditMode = true;
  //   this.conceptId  = item.conceptId;
  //   this.loadConcept(item.conceptId);
  // }
  onSelectLatestConcept(item: LatestConceptItem): void {
    // MUST route on the stable anchor (ConceptId), never the version/display
    // id (CurrentConceptId). Routing on CurrentConceptId sends a display id
    // back to the server as concept_id on the next Update; the backend won't
    // recognize it as an existing anchor and will insert a brand-new
    // ConceptKeys/Concepts row instead of updating the original one.
    if (this.isActiveConcept(item)) return;
    this.router.navigate(['/concept-create', item.ConceptId]);
  }

  /** Whether a Latest-Updates card is the concept currently open on the
   *  page. Compares as STRINGS on purpose: this.conceptId always comes
   *  from the route param (Angular route params are always strings), but
   *  item.ConceptId's *actual* runtime type depends on whatever
   *  /api/latest-updates serializes it as — the `ConceptId: string`
   *  interface field above is only a compile-time annotation, not a
   *  guarantee. If the backend returns it as a JSON number for some rows
   *  (mixed int/varchar concept-id columns, legacy vs new data, etc.), a
   *  strict `===` against the route's string id fails ONLY for those rows
   *  — which is exactly why the active highlight used to work for some
   *  concepts and not others depending on which one you opened.
   *
   *  ALSO checks CurrentConceptId (the display/version id) as a fallback.
   *  This must always route on ConceptId (the stable anchor) — see the
   *  routing note on onSelectLatestConcept() — but if some OTHER entry
   *  point (e.g. the Dashboard's grid) ever navigates here using a
   *  display id instead of the anchor, this keeps the sidebar highlight
   *  working anyway rather than silently failing to match at all. This
   *  is a UI safety net only; it does not fix (and should not be relied
   *  on to mask) an upstream caller sending the wrong id — see that
   *  comment for why sending CurrentConceptId as concept_id on a
   *  subsequent Update is a real data-integrity risk, not just a display
   *  quirk. */
  isActiveConcept(item: LatestConceptItem): boolean {
    const current = String(this.conceptId);
    return String(item.ConceptId) === current || String(item.CurrentConceptId) === current;
  }

  loadUserFiles(): void {
    const user_id = Number(sessionStorage.getItem('userId'));
    if (!user_id) {
      return;
    }

    this.service.getuserfiles(user_id).subscribe({
      next: (res) => {
        console.log('User Files:', res);

        // Populate supporting documents here
        // this.supportingDocs = res.data;
      },
      error: (err) => {
        console.error('Failed to load user files', err);
        this.toastr.error('Failed to load supporting documents', 'Error');
      }
    });
  }

  // ── Activity History ─────────────────────────────────────────────────
  /** Rows actually rendered in the collapsed card — full list once
   *  "View full history" has been clicked, otherwise just the most
   *  recent activityHistoryPreviewCount entries. activityHistory itself
   *  is already newest-first (see concept_history.py), so no extra
   *  sort/slice-from-the-end logic is needed here. */
  get visibleActivityHistory(): ActivityHistoryItem[] {
    return this.showFullActivityHistory
      ? this.activityHistory
      : this.activityHistory.slice(0, this.activityHistoryPreviewCount);
  }

  toggleFullActivityHistory(): void {
    this.showFullActivityHistory = !this.showFullActivityHistory;
  }

  /** Fetches the audit trail for the given concept. Safe to call with an
   *  empty/undefined id (e.g. a brand-new, not-yet-saved concept) — just
   *  clears the list rather than firing a request that can't succeed. */
  loadActivityHistory(id: string): void {
    if (!id) {
      this.activityHistory = [];
      return;
    }
    this.activityHistoryLoading = true;
    this.service.getConceptHistory(id).subscribe({
      next: (res: { success: boolean; history: ActivityHistoryItem[]; dot_color_map?: Record<string, string> }) => {
        this.activityHistory = res.history ?? [];
        this.activityHistoryLoading = false;
        if (res.dot_color_map) {
          this.activityDotLegend = this.buildActivityDotLegend(res.dot_color_map);
        }
      },
      error: (err) => {
        console.error('Failed to load activity history', err);
        this.activityHistory = [];
        this.activityHistoryLoading = false;
      }
    });
  }

  /** Turns the backend's { activity_type: color } map into display-ready
   *  legend rows, merging activity types that share the same color (e.g.
   *  Attachment/Supporting Document both map to yellow) into a single
   *  row so the legend doesn't repeat itself, and inserting spaces into
   *  camelCase types (FieldUpdate -> "Field Update"). */
  private buildActivityDotLegend(map: Record<string, string>): { color: string; colorLabel: string; label: string }[] {
    const labelsByColor = new Map<string, string[]>();
    Object.entries(map).forEach(([activityType, color]) => {
      const label = activityType.replace(/([a-z])([A-Z])/g, '$1 $2');
      const existing = labelsByColor.get(color) ?? [];
      if (!existing.includes(label)) existing.push(label);
      labelsByColor.set(color, existing);
    });
    return Array.from(labelsByColor.entries()).map(([color, labels]) => ({
      color,
      colorLabel: color.charAt(0).toUpperCase() + color.slice(1),
      label: labels.join(' / ')
    }));
  }

  /** Recomputes the popover's fixed-position coordinates from the info
   *  icon's current on-screen position. Called right before the popover
   *  opens (on hover-in and on click) so it always anchors correctly, even
   *  if the page has scrolled or the layout has changed since last time. */
  private positionActivityLegend(): void {
    const wrapEl = this.activityLegendWrapRef?.nativeElement;
    if (!wrapEl) return;

    const rect = wrapEl.getBoundingClientRect();
    const POPOVER_WIDTH = 240;
    const GAP = 6;
    const VIEWPORT_MARGIN = 8;

    // Right-align the popover's right edge with the icon's right edge (like
    // the old `right: 0` did), but clamp so it never runs off the left edge
    // of the viewport.
    let left = rect.right - POPOVER_WIDTH;
    left = Math.max(VIEWPORT_MARGIN, left);

    this.activityLegendPos = { top: rect.bottom + GAP, left };
  }

  /** Hover-in handler on the icon wrapper — repositions the popover so it's
   *  correctly placed for the CSS-driven :hover preview. Also clears any
   *  lingering force-closed state from a previous close, so hovering the
   *  icon again always shows the popover as expected. */
  onActivityLegendHoverStart(): void {
    this.positionActivityLegend();
    this.legendHoverSuppressed = false;
  }

  toggleActivityLegend(): void {
    this.positionActivityLegend();
    this.showActivityLegend = !this.showActivityLegend;
    this.legendHoverSuppressed = false;
  }

  /** Bound to .legend-force-closed — see the matching CSS comment. The
   *  pointer is typically still over the popover when this runs (it's
   *  what you just clicked), so plain :hover on .activity-legend-wrap
   *  would otherwise keep the popover visible until the mouse physically
   *  left it. This flag forces it hidden immediately on click, and is
   *  cleared by the wrap's (mouseleave) once the pointer actually moves
   *  off, so a fresh hover behaves normally again. */
  legendHoverSuppressed = false;

  closeActivityLegend(): void {
    this.showActivityLegend = false;
    this.legendHoverSuppressed = true;
  }

  openActivityDetail(item: ActivityHistoryItem): void {
    this.selectedActivity = item;
  }

  closeActivityDetail(): void {
    this.selectedActivity = null;
  }

  /** Old/new values for date fields (QA Schedule, Production Schedule, etc.)
   *  come back from the API as full timestamps, e.g. "2026-09-11 00:00:00".
   *  The audit popup only needs the date, shown as MM/DD/YYYY. Anything
   *  that isn't a plain date/timestamp string (e.g. a text or dropdown
   *  change) is returned unchanged. */
  formatActivityValue(value: string): string {
    if (!value) {
      return value;
    }
    const match = value.trim().match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T]\d{2}:\d{2}:\d{2})?$/);
    if (!match) {
      return value;
    }
    const [, year, month, day] = match;
    return `${month}/${day}/${year}`;
  }

  private static validDateRange(control: import('@angular/forms').AbstractControl) {
    const val: string = control.value;
    if (!val) return null;
    const year = new Date(val).getFullYear();
    if (isNaN(year) || year < 1900 || year > 3000) {
      return { invalidDateRange: true };
    }
    return null;
  }

  /** Rejects any date before today — backs up the [min] attribute on the
   *  date input, since native min/max can be bypassed by manual keyboard
   *  entry in some browsers (notably Firefox). Compares by calendar day,
   *  not time-of-day, so "today" itself is always valid regardless of
   *  current time.
   *
   *  Only enforced while the control is `dirty` (the user is actively
   *  picking/typing a date right now). patchForm() loads an existing
   *  concept's saved QASchedule/ProductionSchedule via patchValue(),
   *  which leaves the control pristine — so a concept created weeks ago
   *  with a QA/Production date that has since arrived (or passed) must
   *  NOT be flagged here. Those dates recording when QA/production
   *  actually happened are supposed to end up in the past; that's normal,
   *  not an error. Without the dirty check, that concept would become
   *  permanently un-updatable — this validator would block saving ANY
   *  field, forever, just because time moved on since it was created.
   *  Once the user actually edits one of these fields to a new value,
   *  the control becomes dirty and the "no past dates" rule correctly
   *  applies to that new pick. */
  private static notPastDate(control: import('@angular/forms').AbstractControl) {
    const val: string = control.value;
    if (!val) return null;
    if (!control.dirty) return null;
    const selected = new Date(val);
    selected.setHours(0, 0, 0, 0);
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    if (selected < today) {
      return { pastDate: true };
    }
    return null;
  }

  limitToDigits(event: Event, controlName: string, maxDigits: number = 9): void {
    const input = event.target as HTMLInputElement;
    let value = input.value.replace(/\D/g, ''); // strip non-digits

    if (value.length > maxDigits) {
      value = value.slice(0, maxDigits);
    }

    input.value = value;
    this.form.get(controlName)?.setValue(value ? Number(value) : null, { emitEvent: false });
  }

  /** Same digit-only stripping as limitToDigits, but Halo Number specifically
   *  must never be exactly zero. Leading zeros are otherwise left alone —
   *  "0025" stays "0025" — this only blocks the value from being the
   *  single digit "0" itself. */
  limitToDigitsNoZero(event: Event, controlName: string, maxDigits: number = 9): void {
    const input = event.target as HTMLInputElement;
    let value = input.value.replace(/\D/g, ''); // strip non-digits

    // Only reject the exact value "0" — everything else (including
    // values with leading zeros like "0025") passes through untouched.
    if (value === '0') {
      value = '';
    }

    if (value.length > maxDigits) {
      value = value.slice(0, maxDigits);
    }

    input.value = value;
    this.form.get(controlName)?.setValue(value ? Number(value) : null, { emitEvent: false });
  }

  /** Strips anything that isn't a letter or digit as the user types —
   *  Previous Report ID must not contain spaces, punctuation, or any
   *  other special characters. */
  restrictSpecialChars(event: Event, controlName: string): void {
    const input = event.target as HTMLInputElement;
    const value = input.value.replace(/[^a-zA-Z0-9]/g, '');

    input.value = value;
    this.form.get(controlName)?.setValue(value, { emitEvent: false });
  }
  get draftLockMessage(): string {
    return !this.conceptId
      ? 'This concept hasn\'t been saved yet. Save the Concept Information first to enable Client Approval and Supporting Document.'
      : 'This concept is saved as a draft. Save the Concept Information to enable Client Approval and Supporting Document.';
  }
}