# Phase 0b — Windows 11 COM test VM build guide

Goal: a Windows 11 KVM guest that (a) boots straight into a logged-in interactive desktop
(**autologon**), (b) has the **desktop OneNote** signed into M365 with a synced notebook, and
(c) can run `OneNote.Application` COM via pywin32. This is the test/dev environment only —
production runs on each employee's own PC (SPEC §3, §8).

**Acceptance (the one thing that must work at the end):** in the autologon interactive
session, `win32com.client.Dispatch("OneNote.Application").GetHierarchy("", 2)` returns notebook
XML. That single success unblocks Phase 1 onward.

Why so much ceremony (SPEC §2.4): **COM only works in the interactive logged-on session.** A
plain `ssh guest <command>` lands in a non-interactive session where COM against OneNote fails.
So we autologon (always have an interactive session) and later trigger work *into* that session.

---

## Ordered checklist (detail below)

1. [ ] Host: enable KVM, install libvirt + virt tools + swtpm + OVMF
2. [ ] Download Windows 11 ISO + virtio-win ISO
3. [ ] `virt-install` the VM (UEFI + TPM 2.0 + virtio, NAT network)
4. [ ] Install Windows 11 (load virtio storage driver during setup)
5. [ ] Guest: virtio guest agent, Windows Update
6. [ ] Guest: install **desktop OneNote** (M365), sign in, sync a notebook, keep synced
7. [ ] Guest: Python 3.12 + uv + pywin32 (run pywin32 post-install)
8. [ ] Guest: OpenSSH Server (autostart + host key)
9. [ ] Guest: **autologon**
10. [ ] **COM smoke test** in the autologon session ← acceptance gate
11. [ ] Clone repo + `uv sync` + `uv run pytest` on guest
12. [ ] Task Scheduler "run only when logged on" task for Tier 2
13. [ ] Snapshot the clean baseline
14. [ ] Send me the connection info + COM smoke result

---

## 1. Host prerequisites (Ubuntu 24.04)

```bash
# Virtualization is on? (>0 expected; W-1290 has VT-x. Also enable VT-x/VT-d in BIOS if 0.)
egrep -c '(vmx|svm)' /proc/cpuinfo

sudo apt update
sudo apt install -y qemu-kvm libvirt-daemon-system libvirt-clients \
  virtinst virt-manager virt-viewer ovmf swtpm swtpm-tools cpu-checker

kvm-ok                      # should say "KVM acceleration can be used"
sudo usermod -aG libvirt,kvm "$USER"   # then log out/in (or `newgrp libvirt`)
sudo systemctl enable --now libvirtd
virsh list --all           # sanity: talks to libvirt
```

## 2. Download the bits

- **Windows 11 ISO** — https://www.microsoft.com/software-download/windows11
- **virtio-win ISO** (paravirtualized storage/net drivers, big perf win) —
  https://fedorapeople.org/groups/virt/virtio-win/direct-downloads/stable-virtio/virtio-win.iso

Put both somewhere libvirt can read, e.g. `/var/lib/libvirt/images/`.

## 3. Create the VM

Win11 **requires UEFI + Secure Boot + TPM 2.0** — provided here by OVMF + swtpm. Thin qcow2 so
snapshots are cheap (SPEC §2.2: ~120GB, 4 vCPU / 8GB to start; bump RAM if OneNote sync is slow —
host has ~24GB).

```bash
virt-install \
  --name win11-onenote \
  --osinfo win11 \
  --vcpus 4 \
  --memory 8192 \
  --cpu host-passthrough \
  --machine q35 \
  --boot uefi \
  --tpm backend.type=emulator,backend.version=2.0,model=tpm-crd \
  --disk path=/var/lib/libvirt/images/win11-onenote.qcow2,size=120,format=qcow2,bus=virtio \
  --cdrom /var/lib/libvirt/images/Win11.iso \
  --disk path=/var/lib/libvirt/images/virtio-win.iso,device=cdrom \
  --network network=default,model=virtio \
  --graphics spice --video qxl --channel spicevmc
```

Notes:
- If `--osinfo win11` errors (old osinfo-db), use `--os-variant win11` or `--osinfo detect=on,name=win11`.
- `network=default` is libvirt NAT: gives the guest internet (needed for M365/OneNote sync) and a
  `192.168.122.x` IP the host can SSH to. No port-forwarding needed for host→guest.

## 4. Install Windows 11 (via the SPICE console)

The `virt-install` above opens `virt-viewer`. To reconnect later (incl. **remotely over your
existing SSH**, SPEC §2.3):

```bash
# from your laptop, against the Linux host:
virt-viewer --connect qemu+ssh://USER@LINUX_HOST/system win11-onenote
# or: virt-manager --connect qemu+ssh://USER@LINUX_HOST/system
```

During Windows setup:
- **⚠ The disk won't appear** (virtio). Click **Load driver** → browse the virtio-win CD →
  `vioscsi\w11\amd64` (or `viostor\w11\amd64`) → load → the disk shows up.
- TPM/Secure Boot checks pass thanks to swtpm + OVMF.
- A local account is fine for Windows itself; you'll sign into **OneNote** with M365 separately.
  (To skip the MS-account requirement on Win11 Home/Pro: at the network step `Shift+F10` →
  `OOBE\BYPASSNRO` , reboot, choose "I don't have internet" — optional.)

## 5. Guest: drivers + updates

- Open the virtio-win CD in the guest → run `virtio-win-guest-tools.exe` (installs NIC driver +
  qemu-guest-agent → lets the host query the guest IP).
- Run Windows Update fully.

## 6. Guest: OneNote (the COM-critical step)

- Install the **desktop OneNote** from Microsoft 365 (the unified "OneNote" app). 
  **⚠ Not** the old "OneNote for Windows 10" UWP — that one does **not** expose the
  `OneNote.Application` COM interface. Verify in PowerShell:
  ```powershell
  New-Object -ComObject OneNote.Application   # must succeed (no error)
  ```
- Sign into your M365 account, open/create a notebook on OneDrive, let it **sync**, and leave
  sync on. COM reads/writes this live session.

## 7. Guest: Python toolchain

- Install **Python 3.12+** (python.org; tick "Add to PATH").
- Install uv: in PowerShell `powershell -c "irm https://astral.sh/uv/install.ps1 | iex"`.
- pywin32 (the post-install step matters):
  ```powershell
  py -m pip install pywin32
  py -m pywin32_postinstall -install
  ```

## 8. Guest: OpenSSH Server (for the host→guest trigger channel)

```powershell
# elevated PowerShell
Add-WindowsCapability -Online -Name OpenSSH.Server~~~~0.0.1.0
Set-Service sshd -StartupType Automatic
Start-Service sshd
New-NetFirewallRule -Name sshd -DisplayName 'OpenSSH Server' -Enabled True `
  -Direction Inbound -Protocol TCP -Action Allow -LocalPort 22
```
Add your host's public key. **⚠ Gotcha:** for an **admin** user, OpenSSH on Windows reads
`C:\ProgramData\ssh\administrators_authorized_keys` (not `~\.ssh\authorized_keys`), and that file
must be owned/readable only by Administrators+SYSTEM. From the Linux host test:
`ssh USER@GUEST_IP` (find the IP with `virsh domifaddr win11-onenote --source agent`).

## 9. Guest: autologon (the key enabler)

Make the guest boot into a logged-in interactive desktop, so an interactive COM session always
exists:
- Easiest: Sysinternals **Autologon64.exe** (https://learn.microsoft.com/sysinternals/downloads/autologon)
  — enter user + password, click Enable.
- Or `netplwiz` → untick "Users must enter a user name and password".

Reboot and confirm it lands on the desktop with no prompt.

## 10. COM smoke test — the acceptance gate

In the **autologon interactive session** (the SPICE console, not over plain SSH), run:

```python
# save as com_smoke.py, run:  py com_smoke.py
import win32com.client

# Try early binding first (what Win32ComBackend uses):
try:
    app = win32com.client.gencache.EnsureDispatch("OneNote.Application")
    mode = "EnsureDispatch (early binding)"
except Exception as e:
    print("EnsureDispatch failed:", e)
    app = win32com.client.Dispatch("OneNote.Application")
    mode = "Dispatch (late binding)"

xml = app.GetHierarchy("", 2)   # 2 = hsNotebooks
print("MODE:", mode)
print("TYPE of GetHierarchy return:", type(xml))
print(xml[:800])
```

✅ If this prints notebook XML, Phase 0 acceptance is met. **Record which `mode` worked and what
`GetHierarchy` returned** — that resolves the one open question in `Win32ComBackend` (whether the
`[out]` param comes back as the return value, and whether early or late binding to use).

## 11. Run the repo's Tier 1 on the guest

```powershell
git clone git@github.com:chrislin8848/onenote-com-mcp.git C:\onenote-com-mcp
cd C:\onenote-com-mcp
uv sync
uv run pytest          # Tier 1 should be green on Windows too (windows-marked tests now run)
```

## 12. Tier 2 trigger task (SPEC §2.4)

Create the "run only when logged on" scheduled task once (so triggers land in the autologon
session, where COM works):

```powershell
schtasks /create /tn onenote-tier2 /sc once /st 00:00 /it /tr ^
  "cmd /c cd /d C:\onenote-com-mcp && uv run pytest -m windows --junitxml=test-results\tier2.xml -ra > test-results\tier2.log 2>&1"
```
From the Linux host you then drive it with `scripts/remote_test.sh` (set `GUEST_HOST`,
`GUEST_USER`; it rsyncs code, `schtasks /run`s the task, and pulls results + fixtures back).

## 13. Snapshot the clean baseline

Once OneNote is synced and the COM smoke passes:
```bash
virsh snapshot-create-as win11-onenote clean-baseline "OneNote signed in + COM verified"
# revert anytime after destructive tests:  virsh snapshot-revert win11-onenote clean-baseline
```

## 14. Bring these back to me (so I can finish wiring Phase 1+)

- guest IP + SSH user/port, repo path on guest
- the COM smoke result: which binding `mode` worked + the `type()` and first lines of
  `GetHierarchy` output
- 2–3 real page IDs (from the printed hierarchy XML, the `ID="..."` of pages that have:
  mixed text styling, a table, an image) → I'll feed them to `scripts/dump_fixtures.py`

With that I'll finalize `Win32ComBackend` out-param handling, `dump_fixtures.py`, and start
Phase 1 parse/build against the real fixtures.
