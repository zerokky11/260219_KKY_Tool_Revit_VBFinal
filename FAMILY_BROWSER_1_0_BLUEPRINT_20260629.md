# Family Browser 1.0 Blueprint - 2026-06-29

## Product Target

Family Browser 1.0 is not just a family loader. It is a Revit standard-library control surface:

- show only approved families and system types for the active trade;
- compare the current project against the registered standard RVT;
- load, update, or apply standard content through governed workflows;
- keep project open/startup lightweight by using cached snapshots first;
- let modelers request missing content while admins manage standards, permissions, and guard targets;
- keep all shared operational data under the managed folder/homepage source of truth.

## 1.0 Navigation

Use five primary areas instead of the current scattered tab set:

1. **Library**
   - Family Load
   - System Type Load
   - approved list filters
   - selected item inspector with preview, parameters, composition, and status

2. **Model Check**
   - current model scan/compare
   - fingerprint differences
   - missing/unregistered loadable families
   - missing/unregistered system types
   - admin apply queue

3. **Requests**
   - submit missing/exception content request
   - view request queue
   - approve/reject/close
   - attachments and reason capture

4. **Governance**
   - permission snapshot
   - FileGuard target list
   - Excel permission diagnostics
   - role/admin mode state

5. **Standards**
   - registered standard RVT
   - visible standard list Excel/JSON
   - standard RVT manager
   - precise scan / fast stamp check / selected refresh
   - operational readiness and logs

## UI Layout

### Global Shell

- Fixed top context bar:
  - active project;
  - trade / standard slot;
  - role and Admin Mode;
  - readiness status;
  - homepage managed-folder state.
- Left vertical navigation rail with grouped icons/text:
  - Library, Model Check, Requests, Governance, Standards.
- Main work area:
  - task-specific toolbar;
  - dense tables with sticky headers;
  - zero marketing/hero content.
- Right inspector:
  - only visible for Library and Model Check item selection;
  - shows preview, parameters, fingerprints, nested family composition, and next action.

### Library Screen

Layout:

- left tree: trade, group, category;
- center list: approved standard rows only;
- top controls: search, family/system switch, status chips, mismatch filter;
- right inspector: preview, composition, types, parameters, status, actions.

Behavior:

- cache-only on startup;
- nested approved child families hidden from load list;
- composite detail lists nested family names only;
- no live scan unless the user presses an explicit scan/check action.

### Model Check Screen

Layout:

- summary cards: missing, update available, fingerprint different, protected deletion;
- two review tables: loadable families and system types;
- right inspector: difference explanation and admin action plan.

Behavior:

- scan is explicit;
- result dialog should be HTML-style and explain what changed, what was skipped, and what needs admin review.

### Requests Screen

Layout:

- split queue: open requests, mine, needs approval, closed;
- request composer drawer/modal;
- item context carried from Library/Model Check.

Behavior:

- modelers create/submit;
- approvers/admins approve/reject/close;
- network/shared request store when configured, otherwise clear diagnostic.

### Governance Screen

Layout:

- current role and effective permissions;
- FileGuard targets table;
- Excel permission lookup result;
- native command group status.

Behavior:

- FileGuard remains target-only;
- Excel rows must override native guard decisions for guarded permissions;
- no policy means no implicit admin.

### Standards Screen

Layout:

- registration state at top;
- standard RVT manager actions;
- visible standard list setup;
- precise scan/fast check/selected refresh;
- diagnostics and log links.

Behavior:

- managed folder/homepage is source of truth;
- local manual path overrides stay out of the normal user flow;
- heavy standard RVT work only runs from explicit admin actions.

## Implementation Strategy

1. Keep existing domain services where they already work:
   - StandardLibraryRegistrationService;
   - FamilyBrowserStandardListService;
   - ProjectStandardComparisonService;
   - LoadableFamilySync* services;
   - SystemType* services;
   - FamilyBrowserRequest* services;
   - FamilyBrowserSecurityPolicyService;
   - FamilyBrowserNativeCommandGuardService.

2. Add a 1.0 dashboard shell layer:
   - new navigation grouping;
   - clearer pane classes;
   - reusable status/action cards;
   - right inspector as first-class UI, not an accidental hidden panel.

3. Avoid wholesale rewrite of `FamilyBrowserDashboardHtmlForm.cs`:
   - create helper renderers where possible;
   - keep changes small and build after every stage;
   - keep 2019-2023, 2025, and 2027 host folders synchronized.

4. Build order:
   - IA shell and visual system;
   - Library screen polish;
   - Model Check screen polish;
   - Governance/Standards consolidation;
   - request workflow polish;
   - result dialogs.

## Visual Direction

- dense operational tool, not a landing page;
- calm neutral surface with restrained green/teal accents;
- compact typography for Revit modeless-window use;
- obvious next actions;
- tables are primary, cards only for summaries and repeated review items;
- no oversized decorative hero blocks.
