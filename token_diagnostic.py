"""
Token Diagnostic Script
Helps diagnose issues with environment variables and tokens.
"""

import os
import sys
from dotenv import load_dotenv

# Load environment variables
load_dotenv()


def diagnose_token(var_name: str) -> None:
    """Diagnose a specific environment variable."""
    print(f"\n{'='*70}")
    print(f"Diagnosing: {var_name}")
    print('='*70)

    value = os.getenv(var_name)

    if value is None:
        print(f"❌ {var_name} is NOT SET in environment")
        return

    print(f"✅ {var_name} is set")
    print(f"   Type: {type(value)}")
    print(f"   Length: {len(str(value))} characters")

    # Check if it's bytes
    if isinstance(value, bytes):
        print(f"   ⚠️  WARNING: Value is bytes, not string!")
        print(f"   Raw bytes: {value}")
        try:
            decoded = value.decode('utf-8')
            print(f"   Decoded: {decoded[:20]}...")
        except:
            print(f"   ❌ Cannot decode bytes to UTF-8")
    else:
        # Show first and last few characters
        value_str = str(value)
        if len(value_str) > 40:
            display = f"{value_str[:20]}...{value_str[-20:]}"
        else:
            display = value_str
        print(f"   Value: {display}")

    # Check for common issues
    if value == '***':
        print(f"   ❌ PROBLEM: Token is masked as '***'")
        print(f"   This happens in CI/CD when secrets aren't properly injected")

    if value.startswith('***'):
        print(f"   ❌ PROBLEM: Token starts with '***' (partially masked)")

    # Check for whitespace
    value_str = str(value)
    if value_str != value_str.strip():
        print(f"   ⚠️  WARNING: Token has leading/trailing whitespace")
        print(f"   Stripped length: {len(value_str.strip())}")

    # Check for quotes
    if value_str.startswith('"') or value_str.startswith("'"):
        print(f"   ⚠️  WARNING: Token starts with a quote character")

    # Check for control characters
    control_chars = [c for c in value_str if ord(c) < 32]
    if control_chars:
        print(f"   ⚠️  WARNING: Token contains {len(control_chars)} control characters")

    # Check minimum length
    if len(value_str) < 50:
        print(f"   ⚠️  WARNING: Token seems short (usually >50 characters for JWT)")


def main():
    """Run diagnostics."""
    print("\n" + "="*70)
    print("CSFA TOKEN DIAGNOSTIC TOOL")
    print("="*70)

    # Check if .env file exists
    if os.path.exists('.env'):
        print("✅ .env file found")
    else:
        print("⚠️  .env file not found in current directory")

    # Diagnose each required token
    tokens = [
        "ACCESS_TOKEN",
        "LARAVEL_TOKEN",
        "SAT_SESSION",
        "XSRF_TOKEN"
    ]

    for token in tokens:
        diagnose_token(token)

    print("\n" + "="*70)
    print("ENVIRONMENT INFORMATION")
    print("="*70)
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Environment type: {'CI/CD' if os.getenv('CI') else 'Local'}")

    if os.getenv('GITHUB_ACTIONS'):
        print(f"GitHub Actions: Yes")
        print(f"Runner OS: {os.getenv('RUNNER_OS', 'Unknown')}")

    print("\n" + "="*70)
    print("RECOMMENDATIONS")
    print("="*70)

    all_set = all(os.getenv(t) for t in tokens)
    any_masked = any(os.getenv(t) in ['***', None] or (os.getenv(t) or '').startswith('***') for t in tokens)

    if not all_set:
        print("❌ Some tokens are missing:")
        print("   1. Check your .env file exists and is in the correct location")
        print("   2. Verify variable names match exactly (case-sensitive)")
        print("   3. Ensure there are no typos in variable names")

    if any_masked:
        print("❌ Some tokens appear to be masked:")
        print("   For GitHub Actions:")
        print("   1. Go to Settings > Secrets and variables > Actions")
        print("   2. Add each token as a repository secret")
        print("   3. In your workflow, use: ${{ secrets.ACCESS_TOKEN }}")
        print("   4. Make sure secrets are passed to the job environment")

    print("\n✅ Diagnostic complete!")
    print("="*70 + "\n")


if __name__ == "__main__":
    main()
