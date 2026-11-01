@echo off
rem Chay tram SAP (RemoteSigned: chay duoc script cuc bo; install.ps1 da Unblock-File cac file tai ve).
rem Neu cong ty chan, hoi IT cho phep thu muc C:\TramSAP thay vi doi sang Bypass.
rem Chay tram SAP thu nho duoi thanh tac vu (khong chiem man hinh). Mo lai cua so tu thanh tac vu de xem nhat ky; dong cua so do de dung tram.
cd /d "%~dp0"
start "Tram SAP - PR truoc" /min powershell -NoProfile -ExecutionPolicy RemoteSigned -File "%~dp0TramSAP.ps1"
