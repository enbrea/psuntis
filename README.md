# PsUntis

[![PowerShell Gallery - PsUntis](https://img.shields.io/badge/PowerShell%20Gallery-PsUntis-blue.svg)](https://www.powershellgallery.com/packages/PsUntis)
[![Minimum Supported PowerShell Version](https://img.shields.io/badge/PowerShell-7-blue.svg)](https://github.com/enbrea/psuntis)
[![Minimum Supported PowerShell Version](https://img.shields.io/badge/PowerShell-5.1-blue.svg)](https://github.com/enbrea/psuntis)

## Einführung

PsUntis ist ein [PowerShell-Modul](https://www.powershellgallery.com/packages/PsUntis) für die Nutzung von Untis auf Kommandozeilenebene. 

+ Das Cmdlet `Start-UntisExport` kapselt den Untis-Kommandozeilenbefehl für Datei-Exporte (GPU oder XML) und implementiert darüber hinaus folgende zusätzliche Annehmlichkeiten:

	+ GPU-Dateien (in jeder beliebigen Kombination) und die XML-Datei können auf Wunsch mit einem einzigen Aufruf erzeugt werden. 
	
	+ Der Export wird stets auf einer Kopie der Untis-Daten durchgeführt, das erlaubt ein transaktionelles Verhalten beim Export mehrerer Dateien.
	
	+ Werden die Exportdateien nicht erstellt, gibt das Cmdlet eine Fehlermeldung aus.

+ Das Cmdlet `Start-UntisBackup` kapselt den Untis-Kommandozeilenbefehl für Datei-Backups (nur unter Untis MultiUser verfügbar) und implementiert darüber hinaus folgende zusätzliche Annehmlichkeiten:

	+ Wird die Backupdatei nicht erstellt, gibt das Cmdlet eine Fehlermeldung aus.

+ Die Cmdlets `Get-UntisTerms` und `Set-UntisTermAsActive` erlauben das Auslesen aller definierten Perioden aus einer gpn-Datei bzw. das Ändern der aktiven Periode in einer gpn-Datei.

## PsUntis installieren

Vorgehensweise:

1. Starte PowerShell (Microsoft PowerShell 7 oder PowerShell 5.1)

2. Tippe `Install-Module PsUntis` ein und bestätige.

## PsUntis aktualisieren

Vorgehensweise:

1. Starte PowerShell (Microsoft PowerShell 7 oder PowerShell 5.1)

2. Tippe `Update-Module PsUntis` ein und bestätige.

## Überprüfen, welche Version von PsUntis installiert ist

Vorgehensweise:

1. Starte PowerShell (Microsoft PowerShell 7 oder PowerShell 5.1)

2. Tippe `Get-InstalledModule PsUntis` ein und bestätige.

## Dokumentation

Dokumentation der Cmdlets findest Du im [GitHub-Wiki](https://github.com/enbrea/psuntis/wiki).

## Kann ich mithelfen?

Ja, sehr gerne. Der beste Weg mitzuhelfen ist es, den Quellcode auszuprobieren, Rückmeldung per Issue-Tracker zu geben und/oder eigene Pull-Requests zu generieren. Oder schreibe uns einfach eine E-Mail unter enbrea@stueber.de.
