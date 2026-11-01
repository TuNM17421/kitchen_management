@echo off
rem Chay tram SAP thu nho duoi thanh tac vu (khong chiem man hinh). Mo lai cua so tu thanh tac vu de xem nhat ky; dong cua so do de dung tram.
cd /d "%~dp0"
start "Tram SAP - PR truoc" /min powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0TramSAP.ps1"
