import os
import sys
from dulwich.repo import Repo
from dulwich import porcelain

repo_path = r"c:\Reinforcement Learning of Sports"

print(f"Initializing repository at {repo_path}...")
try:
    repo = Repo(repo_path)
    print("Existing git repository found.")
except Exception:
    repo = Repo.init(repo_path)
    print("New git repository initialized.")

# Stage all tracked and new files
porcelain.add(repo, paths=["."])
print("Files staged successfully.")

# Check status
status = porcelain.status(repo)
print(f"Staged: {len(status.staged['add']) + len(status.staged['modify'])} files")

# Commit
commit_msg = b"feat: Complete Scopus Q1 MARL football simulation pipeline with 8-figure visual portfolio"
commit_id = porcelain.commit(
    repo,
    message=commit_msg,
    author=b"John Wendy <johnwendycn@users.noreply.github.com>",
    committer=b"John Wendy <johnwendycn@users.noreply.github.com>"
)
print(f"Commit created: {commit_id.decode('ascii') if isinstance(commit_id, bytes) else commit_id}")

# Set main branch
porcelain.branch_create(repo, b"main", force=True)
print("Branch 'main' set.")
