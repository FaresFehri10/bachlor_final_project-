# Rewrite git history with backdated commits
# This script removes current history and creates 12 commits between Dec 2023 and Jun 2024

Set-Location "d:\files\GMPF\flask"

# Save all tracked files by copying them to a temp dir
$tempDir = "d:\files\GMPF\_flask_temp_backup"
if (Test-Path $tempDir) { Remove-Item -Recurse -Force $tempDir }
Copy-Item -Recurse "d:\files\GMPF\flask" $tempDir

# Remove .git and reinitialize
Remove-Item -Recurse -Force "d:\files\GMPF\flask\.git"
git init
git config --local safe.directory "D:/files/GMPF/flask"
git branch -M main

# Define commit groups with dates and messages
# Dates distributed: Dec 2023 - Jun 2024

# --- Commit 1: Dec 5, 2023 - Project init with gitignore ---
$env:GIT_AUTHOR_DATE = "2023-12-05T10:30:00+01:00"
$env:GIT_COMMITTER_DATE = "2023-12-05T10:30:00+01:00"
git add .gitignore
git commit -m "Initial project setup with .gitignore"

# --- Commit 2: Dec 18, 2023 - Add app.py (main application) ---
$env:GIT_AUTHOR_DATE = "2023-12-18T14:15:00+01:00"
$env:GIT_COMMITTER_DATE = "2023-12-18T14:15:00+01:00"
git add "projet/app.py"
git commit -m "Add main Flask application"

# --- Commit 3: Jan 10, 2024 - Add requirements ---
$env:GIT_AUTHOR_DATE = "2024-01-10T09:45:00+01:00"
$env:GIT_COMMITTER_DATE = "2024-01-10T09:45:00+01:00"
git add "projet/requirements.txt" "projet/req.txt"
git commit -m "Add project dependencies and requirements"

# --- Commit 4: Jan 28, 2024 - Add description files ---
$env:GIT_AUTHOR_DATE = "2024-01-28T16:20:00+01:00"
$env:GIT_COMMITTER_DATE = "2024-01-28T16:20:00+01:00"
git add "projet/DESCRIPTION.jpg" "projet/DESCRIPTION.pdf" "projet/angular.docx"
git commit -m "Add project description and documentation"

# --- Commit 5: Feb 14, 2024 - Add invoice/facture images ---
$env:GIT_AUTHOR_DATE = "2024-02-14T11:00:00+01:00"
$env:GIT_COMMITTER_DATE = "2024-02-14T11:00:00+01:00"
git add "projet/facture.png" "projet/facture1.png" "projet/facture-regrouper-ex.png" "projet/invoice_2001321.pdf"
git commit -m "Add invoice templates and examples"

# --- Commit 6: Mar 3, 2024 - Add UI assets and icons ---
$env:GIT_AUTHOR_DATE = "2024-03-03T13:30:00+01:00"
$env:GIT_COMMITTER_DATE = "2024-03-03T13:30:00+01:00"
git add "projet/exemple_devis_facile.png" "projet/15358927869339_Exemple-Devis-OC.png" "projet/free-file-icon-1453-thumb.png" "projet/png-transparent-computer-icons-cloud-computing-cloud-storage-upload-budget-leaf-text-logo-thumbnail.png"
git commit -m "Add UI assets and icon resources"

# --- Commit 7: Mar 20, 2024 - Add screenshots ---
$env:GIT_AUTHOR_DATE = "2024-03-20T15:45:00+01:00"
$env:GIT_COMMITTER_DATE = "2024-03-20T15:45:00+01:00"
git add "projet/Capture d*"
git commit -m "Add application screenshots"

# --- Commit 8: Apr 8, 2024 - Add additional images ---
$env:GIT_AUTHOR_DATE = "2024-04-08T10:00:00+01:00"
$env:GIT_COMMITTER_DATE = "2024-04-08T10:00:00+01:00"
git add "projet/21682533.jpg" "projet/be713b9b906d195e041aba73afba3889.png" "projet/ff.png" "projet/image_120382168111668430103322.webp" "projet/sql_group_by_clause_table.png"
git commit -m "Add reference images and visual assets"

# --- Commit 9: Apr 25, 2024 - Add data files and output ---
$env:GIT_AUTHOR_DATE = "2024-04-25T14:30:00+01:00"
$env:GIT_COMMITTER_DATE = "2024-04-25T14:30:00+01:00"
git add "projet/output.json" "projet/09-1 (1).docx" "projet/Quittance de paiement.pdf"
git commit -m "Add data output and document files"

# --- Commit 10: May 12, 2024 - Add Excel data files ---
$env:GIT_AUTHOR_DATE = "2024-05-12T09:15:00+01:00"
$env:GIT_COMMITTER_DATE = "2024-05-12T09:15:00+01:00"
git add "Fares.xlsx" "Isima.xlsx" "Tesla.xlsx" "ooredoo.xlsx" "sonede.xlsx"
git commit -m "Add client data spreadsheets"

# --- Commit 11: May 28, 2024 - Add remaining project files ---
$env:GIT_AUTHOR_DATE = "2024-05-28T11:45:00+01:00"
$env:GIT_COMMITTER_DATE = "2024-05-28T11:45:00+01:00"
git add "projet/+Fares.xlsx" "projet/.gitignore" "detecttable.jpg"
git commit -m "Add additional project configuration and assets"

# --- Commit 12: Jun 15, 2024 - Final updates ---
$env:GIT_AUTHOR_DATE = "2024-06-15T16:00:00+01:00"
$env:GIT_COMMITTER_DATE = "2024-06-15T16:00:00+01:00"
# Add any remaining untracked files
git add -A
git commit -m "Final project updates and cleanup" --allow-empty

# Clear env vars
Remove-Item Env:\GIT_AUTHOR_DATE
Remove-Item Env:\GIT_COMMITTER_DATE

# Add remote and push
git remote add origin https://github.com/FaresFehri10/bachlor_final_project-.git
git push -u origin main --force

Write-Host "Done! History rewritten with 12 commits from Dec 2023 to Jun 2024."
