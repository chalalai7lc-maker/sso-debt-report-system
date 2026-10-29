# DEPLOY PLAN: SSO_Debt_Report_System (fix/tls-verify @ 246d93f)

Status: **waiting for PM review and Boss approval.** Nothing has been deployed.
Live dir: `/home/ubuntu/projects/SSO_Debt_Report_System` (live HEAD `95dce4c`) · systemd unit: `sso-debt-report.service` · port 8001
Rules: **no git merge/reset/checkout/stash in the live dir.** Files are copied one by one from `git archive 246d93f`. Restart is a separate step that needs approval.
The deploy contains only code commit `246d93f`. This DEPLOY-PLAN.md is not deployed.

## 0. Files to deploy
Modified (existing in live):
```
requirements.txt
sso_client.py
```
Added (new in live):
```
tests/test_tls_verify.py
```

## 1. Dirty-file overlap check (live dir)
Live `git status --porcelain` at plan time:
```
 D Input/CN200090-1_20260914_161817.xlsx
 M data/DATA-EMP-WCF.xlsx
?? Input/CN200090-1_20261027_140652.xlsx
```
None of these is in the deploy list, so there is **no overlap**. They are data files and must stay untouched (do not `git checkout`/`add` them).

Re-run this check right before deploying. It must print `OK no overlap`:
```bash
LIVE=/home/ubuntu/projects/SSO_Debt_Report_System
cat > /tmp/SSO_Debt_Report_System-deploy-files.txt <<'EOF'
requirements.txt
sso_client.py
tests/test_tls_verify.py
EOF
git -C "$LIVE" status --porcelain --ignored | awk '{print $NF}' > /tmp/SSO_Debt_Report_System-dirty.txt
grep -Fxf /tmp/SSO_Debt_Report_System-dirty.txt /tmp/SSO_Debt_Report_System-deploy-files.txt && echo 'STOP: overlap with dirty/ignored file' || echo 'OK no overlap'
# added files must not already exist in live:
while read f; do [ -e "$LIVE/$f" ] && echo "STOP: $f already exists"; done <<'EOF'
tests/test_tls_verify.py
EOF
# modified files must still match the base commit (nobody edited them):
while read f; do [ "$(git -C "$LIVE" rev-parse 95dce4c:"$f")" = "$(git -C "$LIVE" hash-object "$LIVE/$f")" ] || echo "STOP: $f differs from 95dce4c"; done <<'EOF'
requirements.txt
sso_client.py
EOF
```

## 2. Credentials (path only; Boss sets the values)
No new credential. Optional: if SSO Core uses a private CA, add `SSO_CA_BUNDLE=/home/ubuntu/.config/sso-debt/ca.pem` to the existing (git-ignored) `.env`.

## 3. Pre-checks
- The prod venv still has `httpx2` installed (harmless). Removing it from requirements.txt does not uninstall it; leave it, or `pip uninstall httpx2` later as a separate, approved step.
- Before restart, verify that the SSO cert validates: `$LIVE/.venv/bin/python -c "import httpx,sso_client as s;print(httpx.get(s.BASE_URL, verify=True, timeout=15).status_code)"`. Any HTTP code is OK; an SSL error is not. If you get an SSL error, get the CA first and set SSO_CA_BUNDLE.

## 4. Stage (from the R&D clone, not the live repo)
```bash
# run sections 4-6 and 9 in ONE shell session (TS/STAGE/LIVE are reused)
TS=$(date +%Y%m%d-%H%M%S); CLONE=/home/ubuntu/worktrees/SSO_Debt_Report_System-sec; LIVE=/home/ubuntu/projects/SSO_Debt_Report_System
STAGE=/home/ubuntu/staging/SSO_Debt_Report_System-$TS; mkdir -p "$STAGE"
git -C "$CLONE" fetch -q origin fix/tls-verify && test "$(git -C "$CLONE" rev-parse --short=7 246d93f)" = "246d93f"
git -C "$CLONE" archive 246d93f $(tr '\n' ' ' < /tmp/SSO_Debt_Report_System-deploy-files.txt) | tar -x -C "$STAGE"
( cd "$STAGE" && sha256sum $(cat /tmp/SSO_Debt_Report_System-deploy-files.txt) ) > "$STAGE.sha256"
# verify that staging matches git blobs exactly:
while read f; do [ "$(git -C "$CLONE" rev-parse 246d93f:"$f")" = "$(git hash-object "$STAGE/$f")" ] || echo "BAD $f"; done < /tmp/SSO_Debt_Report_System-deploy-files.txt
cd "$STAGE" && $LIVE/.venv/bin/python -m pytest -q -p no:cacheprovider tests/test_tls_verify.py
```

## 5. Backup originals
```bash
mkdir -p ~/archive
tar -czf ~/archive/SSO_Debt_Report_System-predeploy-$TS.tgz -C "$LIVE" requirements.txt sso_client.py
tar -tzf ~/archive/SSO_Debt_Report_System-predeploy-$TS.tgz      # must list every modified file
printf '%s\n' tests/test_tls_verify.py > ~/archive/SSO_Debt_Report_System-added-$TS.txt
```

## 6. Copy one file at a time (atomic rename in the same directory)
```bash
while read f; do
  mkdir -p "$(dirname "$LIVE/$f")"
  cp -p "$STAGE/$f" "$LIVE/$f.deploy-$TS.tmp"
  ( cd "$STAGE" && sha256sum "$f" ) | awk '{print $1}' > /tmp/want
  [ "$(sha256sum "$LIVE/$f.deploy-$TS.tmp" | awk '{print $1}')" = "$(cat /tmp/want)" ] || { echo "SHA MISMATCH $f"; rm -f "$LIVE/$f.deploy-$TS.tmp"; break; }
  mv -f "$LIVE/$f.deploy-$TS.tmp" "$LIVE/$f"     # rename(2) within the same dir = atomic
done < /tmp/SSO_Debt_Report_System-deploy-files.txt
( cd "$LIVE" && sha256sum -c "$STAGE.sha256" )     # all must be OK
find "$LIVE" -name "*.deploy-$TS.tmp"               # must be empty
```
The live git will now show these files as modified or untracked. This is expected because we do not touch git in the live dir.

## 7. Restart (separate step, needs explicit approval)
```bash
sudo systemctl restart sso-debt-report.service && sleep 3 && systemctl is-active sso-debt-report.service
```

## 8. Health checks
```
curl -s -o /dev/null -w '%{http_code}\n' http://127.0.0.1:8001/api/system-info   # expect 200 (this app has no auth yet)
```
Run one real "ตรวจสอบ SSO CORE" check from the UI; expect no SSL/ConnectError in `journalctl -u sso-debt-report -n 50 --no-pager`.

## 9. Rollback
```bash
tar -xzf ~/archive/SSO_Debt_Report_System-predeploy-$TS.tgz -C "$LIVE"                       # restore modified files
while read f; do rm -f "$LIVE/$f"; done < ~/archive/SSO_Debt_Report_System-added-$TS.txt    # delete newly added files
( cd "$LIVE" && git status --porcelain )   # should return to the pre-deploy dirty list in section 1
sudo systemctl restart sso-debt-report.service          # needs approval, same as step 7
```
Credential files in ~/.config can stay; the old code ignores them.
