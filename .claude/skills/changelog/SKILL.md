---
name: changelog
description: Add the changes made since the latest changelog entry to changelog.js as a new release (English and Dutch). Use at the end of a session, or when the user asks to update the changelog or add a release.
---

# Update the changelog

Add everything that changed since the latest entry in `CHANGELOG` (`changelog.js`) as one release. The user commits the result; do not commit.

## 1. Find what is new

1. Read the top entry of `CHANGELOG` in `changelog.js`. Note its `version`.
2. Find the commit that added that version:
   `git log --format=%H -S "version: '<version>'" -- changelog.js | tail -1`
3. Collect the changes after that point:
   - Top entry committed: everything after that commit. Use `git log --oneline <commit>..HEAD`, `git diff <commit>` (commits plus uncommitted work) and `git status --short` for untracked files.
   - Top entry not committed yet (step 2 finds nothing, or `git diff HEAD -- changelog.js` shows it as added): that entry is still open. Compare the working tree with `HEAD` and add only what the open entry does not describe yet. Update that entry; do not create a new version.
4. Read the diffs, not only commit messages. Messages can be vague or incomplete.
5. Ignore changes users do not see: `changelog.js` itself, `README.md`, `CLAUDE.md`, `.claude/`, tooling, pure refactors.

If nothing user-facing changed, tell the user and stop.

## 2. Pick the version

Bump from the top entry's version (semver):

- New feature → minor: 1.3.0 → 1.4.0
- Only fixes, layout or style tweaks, speed → patch: 1.3.0 → 1.3.1
- Breaking for users (feature removed, file format changed) → major: 1.3.0 → 2.0.0

Several kinds of change together: take the biggest bump. An open entry (step 1.3) keeps its version, unless the new work needs a bigger bump; then raise that entry's version.

## 3. Write the entry

Add it at the top of the `CHANGELOG` array, in the same shape as the existing entries:

```js
  {
    version: '1.4.0',
    date: 'YYYY-MM-DD', // today
    changes: {
      en: [
        'One short sentence per change.',
      ],
      nl: [
        'Eén korte zin per wijziging.',
      ],
    },
  },
```

Rules for the text:

- Write for users, not developers: what they can do now or what changed for them. No file names, function names or commit hashes.
- One change per item, one short sentence, ending with a period.
- Most important change first.
- `en` and `nl` hold the same items in the same order. Do not add other languages; they fall back to English.

## 4. Report

Tell the user the version, the items (Dutch), and which commits or uncommitted changes they cover. Remind them the change is not committed.
