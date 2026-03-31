#!/usr/bin/env python3
"""
Verification script for MiniMax CLI MCP integration in OfficeCLI.
Validates source code changes without requiring .NET SDK.
"""
import re
import sys
import os
import json

BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(BASE, "..", "..")
SRC = os.path.join(ROOT, "src", "officecli", "Core")

PASS = 0
FAIL = 0

def check(desc, condition):
    global PASS, FAIL
    if condition:
        PASS += 1
        print(f"  PASS: {desc}")
    else:
        FAIL += 1
        print(f"  FAIL: {desc}")

def read(path):
    with open(os.path.join(ROOT, path), "r") as f:
        return f.read()

print("=" * 60)
print("MiniMax MCP Integration Verification")
print("=" * 60)

# --- McpInstaller.cs ---
print("\n--- McpInstaller.cs ---")
mcp = read("src/officecli/Core/McpInstaller.cs")

check("Install switch has 'minimax' case",
      'case "minimax" or "minimax-cli":' in mcp)

check("Install switch calls InstallMiniMax()",
      "InstallMiniMax();" in mcp)

check("Uninstall switch has 'minimax' case",
      'case "minimax" or "minimax-cli":\n                UninstallJson("MiniMax CLI", GetMiniMaxMcpPath(), "mcpServers");' in mcp)

check("GetMiniMaxMcpPath() method exists",
      "GetMiniMaxMcpPath()" in mcp)

check("GetMiniMaxMcpPath returns .minimax/mcp.json",
      '".minimax", "mcp.json"' in mcp)

check("InstallMiniMax() method exists",
      "private static void InstallMiniMax()" in mcp)

check("InstallMiniMax calls InstallJson with 'MiniMax CLI'",
      'InstallJson("MiniMax CLI", GetMiniMaxMcpPath(), "mcpServers")' in mcp)

check("ListStatus includes MiniMax CLI",
      'CheckJsonStatus("MiniMax CLI", GetMiniMaxMcpPath())' in mcp)

check("Help text mentions minimax in supported list",
      "minimax (MiniMax CLI)" in mcp)

check("Uninstall help mentions minimax",
      "lms, claude, cursor, vscode, minimax" in mcp)

check("Register help mentions minimax",
      "lms, claude, cursor, vscode, minimax" in mcp)

check("MiniMax section has proper header comment",
      "// ==================== MiniMax CLI ====================" in mcp)

# --- Installer.cs ---
print("\n--- Installer.cs ---")
inst = read("src/officecli/Core/Installer.cs")

check("McpTargets includes minimax entry",
      '("minimax", ".minimax"' in inst)

check("McpTargets minimax has correct skill aliases",
      '["minimax", "minimax-cli"]' in inst)

# --- Program.cs ---
print("\n--- Program.cs ---")
prog = read("src/officecli/Program.cs")

check("MCP help text includes minimax",
      "lms, claude, cursor, vscode, minimax" in prog)

# --- README.md ---
print("\n--- README.md ---")
readme = read("README.md")

check("README MCP section lists minimax command",
      "officecli mcp minimax" in readme)

check("README MCP section shows MiniMax CLI comment",
      "# MiniMax CLI" in readme)

# --- README_zh.md ---
print("\n--- README_zh.md ---")
readme_zh = read("README_zh.md")

check("Chinese README MCP section lists minimax command",
      "officecli mcp minimax" in readme_zh)

check("Chinese README MCP section shows MiniMax CLI comment",
      "# MiniMax CLI" in readme_zh)

# --- SkillInstaller.cs (pre-existing, verify still intact) ---
print("\n--- SkillInstaller.cs (pre-existing) ---")
skill = read("src/officecli/Core/SkillInstaller.cs")

check("SkillInstaller still has MiniMax entry",
      '"minimax", "minimax-cli"' in skill)

check("SkillInstaller MiniMax display name is 'MiniMax CLI'",
      '"MiniMax CLI"' in skill)

check("SkillInstaller MiniMax detect dir is '.minimax'",
      '".minimax"' in skill)

# --- Consistency checks ---
print("\n--- Consistency ---")

# Count MiniMax references across Install/Uninstall/ListStatus
install_minimax = mcp.count('case "minimax" or "minimax-cli":')
check("Install and Uninstall both handle minimax (2 switch cases)",
      install_minimax == 2)

# Verify all other targets still work (no regression)
for target in ["claude", "cursor", "vscode", "lms"]:
    check(f"Target '{target}' still present in Install switch",
          f'case "{target}"' in mcp or f'"{target}" or' in mcp)

# --- Summary ---
print("\n" + "=" * 60)
total = PASS + FAIL
print(f"Results: {PASS}/{total} passed, {FAIL} failed")
if FAIL > 0:
    print("VERIFICATION FAILED")
    sys.exit(1)
else:
    print("ALL CHECKS PASSED")
    sys.exit(0)
