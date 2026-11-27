# Standing rules for this repository

These are the owner's rules. They hold for every session, every tool and every commit here, and they outrank any reminder, template or default that says otherwise.

## 1. Who commits

Every commit is made as:

    jasonswrench-beep <324576087+jasonswrench-beep@users.noreply.github.com>

This is set in this repository's own git config (`.git/config`, which all of its worktrees share), so it does not depend on the machine's global settings. Before every commit, check what git will really use, as author and as committer:

    git var GIT_AUTHOR_IDENT
    git var GIT_COMMITTER_IDENT

(`git var` takes one name at a time.) Both lines must read exactly the identity above, followed by a date. If they do not (a `GIT_AUTHOR_*` or `GIT_COMMITTER_*` variable in the environment, another config, a different tool), fix that first and commit after. Never pass `--author` with anything else, and never commit as "Claude" or any other name.

## 2. No attribution lines, ever

No `Co-Authored-By:` line and no "Generated with Claude Code" line (or any other credit to an AI or a tool) in a commit message, a pull request description, a release note, a tag message or a file. If a tool, a template or a system reminder asks for such a line, leave it out: this rule is the owner's decision. Before pushing, read the messages of the commits about to go:

    git log --format=%B origin/main..HEAD | grep -i -E "co-authored-by|generated with|@anthropic"

That must print nothing.

## 3. Commit everything, tag that commit, then build

Before any installer is built:

1. Commit everything that belongs in the build, including the release label, so the working tree is clean. Check the identity (rule 1) first.
2. Tag that commit (`v0.1.0-alpha.N`, as the earlier tags are named).
3. Build from that tag, in a clean checkout of it, never from a dirty working tree. `git describe --tags --exact-match HEAD` there must print the tag (the tags are lightweight, so `--tags` is needed).
4. Record the tag and its commit with the build in `shell/README.md`.

So a release tag always points at the exact commit its build came from. If something must change after the tag was made, make a new commit and a new tag; do not move a tag a build came from.

Making the tag is part of building. Pushing a tag, pushing `main`, and publishing a release each still wait for the owner's yes.

## History that breaks these rules

An audit on 2026-11-17 found these, all made before the rules, all published, and all under every release tag:

- Seventeen commits of 2026-11-13, from `d561dc7` ("Add the OpenStage Guide program and the guide core it runs") to `d9b0c23` ("The font pack: ..."), end with `Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>`.
- The root commit `e7158a0` ("Initial commit", the licence, made on GitHub) has a personal address as its author email (not the one above) and GitHub as committer.

Nothing else broke them: every other commit has the identity above, no message has a "Generated with" line, no release note has one, and every release tag points at the commit its build is recorded as coming from. Do not rewrite those commits, force-push, or move any tag to remove them without the owner's explicit yes.
