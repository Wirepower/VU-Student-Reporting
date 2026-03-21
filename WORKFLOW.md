# Branch Workflow Guide (VS 2026 + Terminal)

This guide is for your current branch model:

- `master` (stable baseline)
- `test/phase1-ui-responsive-dpi`
- `test/phase2-generic-update-messaging`
- `test/phase3-unit-profiling-api-view`
- `test/phase4-preserve-sql-checkbox-behavior`
- `test/phase5-global-exemplar-account-transition`
- `test/phase6-sql-persisted-exemplar-email-override`
- `test/phase7-v3-hardening-wrapup`

---

## 1) Start of each coding session (always do this)

### Visual Studio 2026 (GUI)
2. Checkout `master`.
3. Pull latest (`master`).
4. Checkout your working branch.
5. Merge `master` into your current branch:
   - Right-click your current branch
   - **Merge Into Current Branch...**
   - Select `master`
6. Rebuild solution and run a quick smoke test.

### Terminal
```bash
git checkout master
git pull origin master
git checkout <your-branch>
git merge master
```

---

## 2) While coding

### Visual Studio 2026 (GUI)
1. Make small changes.
2. Build frequently.
3. Open **Git Changes**.
4. Stage files.
5. Commit with clear message.

### Terminal
```bash
git add <files>
git commit -m "Clear descriptive message"
```

---

## 3) Push your branch

### Visual Studio 2026 (GUI)
- In **Git Changes**, click **Push**.

### Terminal
```bash
git push -u origin <your-branch>
```

---

## 4) Merge branch to master (when phase is ready)

### Safe order
1. Ensure branch is up to date with latest `master` (Section 1).
2. Complete full test pass.
3. Merge via Pull Request (recommended).

### Terminal merge option
```bash
git checkout master
git pull origin master
git merge <your-branch>
git push -u origin master
```

---

## 5) Handling merge conflicts

### Visual Studio 2026 (GUI)
1. Open conflict editor.
2. Choose current/incoming/both carefully.
3. Mark resolved.
4. Commit merge.

### Terminal
```bash
git status
# edit conflicted files
git add <resolved-files>
git commit
```

---

## 6) Quick commands you will use most

```bash
# See current branch
git branch --show-current

# See changed files
git status --short

# See all local and remote phase branches
git branch --all --list "*test/phase*"
```

---

## 7) Recommended branch testing order

1. `master` (baseline already in production)
2. `test/phase1-ui-responsive-dpi` (responsive layout and DPI behavior improvements)
3. `test/phase2-generic-update-messaging` (cleaner user-facing update wording)
4. `test/phase3-unit-profiling-api-view` (unit profiling API view flow)
5. `test/phase4-preserve-sql-checkbox-behavior` (ensures profiling view does not alter SQL-driven unit checkboxes)
6. `test/phase5-global-exemplar-account-transition` (admin API settings/token flow for moving from personal to global Exemplar account)
7. `test/phase6-sql-persisted-exemplar-email-override` (store alternate Exemplar email by student in SQL to avoid repeat prompts)
8. `test/phase7-v3-hardening-wrapup` (stability hardening and final polish items)

