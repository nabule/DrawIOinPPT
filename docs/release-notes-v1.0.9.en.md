# DrawioPpt v1.0.9 Release Notes

[English Home](../README.en.md) | [中文](./release-notes-v1.0.9.md)

## Problem Solved First

The package already wrote the new BuildId, but the Word and PowerPoint Ribbons could not display it. Users therefore could not tell whether Office had actually loaded the upgraded add-in.

## What Changed

- Word and PowerPoint now both expose the version callback used by the Office Ribbon, so the info area shows `版本：v1.0.9+<git-short-hash>`.
- The settings footer continues to show the same BuildId.
- Installation verification starts real Word and PowerPoint, reads the version returned by both add-ins, and compares it with installed `BuildInfo.txt`.

## How to Use It

Install `DrawioPpt-v1.0.9.zip`, reopen Word or PowerPoint, and check the version in the `Draw.io` Ribbon information area. It must match the BuildId printed by the installer.

## Validation

- The version-display regression test covers the Ribbon XML, COM entry callbacks, version lookup, and installer verification chain for both hosts.
- Full Office E2E passed `13/13`, and standard local installation returned a BuildId matching `BuildInfo.txt` from both Word and PowerPoint.
