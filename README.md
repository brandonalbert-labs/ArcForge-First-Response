# ArcForge Battlestation Health Toolkit

A lightweight PowerShell-based Windows health-check toolkit for gamers, creators, developers, homelabbers, and IT pros who want to quickly verify that a PC is ready for work, gaming, or production.

## Project Goal

ArcForge Battlestation Health Toolkit provides a quick baseline health check for Windows workstations.

The goal is to help identify common issues related to system health, uptime, storage, networking, installed tools, basic security posture, Windows Update readiness, running processes, and core Windows services.

This project started as a fictional game studio IT toolkit, but the long-term direction is broader: a practical health-check tool for gaming PCs, creator workstations, developer systems, homelab admin machines, and secure workstations.

## Current Features

- System information check
- Uptime check
- Process readiness check
- Top memory process visibility
- Core Windows service readiness check
- Disk space check
- Network connectivity check
- DNS resolution check
- Required software checks
- Firewall status check
- Antivirus provider detection
- Local administrator group review
- Windows Update readiness check
- Pending reboot detection
- Last installed hotfix check
- Console summary output
- TXT report export

## Current Usage

Run from PowerShell:

```powershell
.\scripts\Invoke-BattlestationHealthCheck.ps1