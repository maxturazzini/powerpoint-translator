#!/usr/bin/env python3
"""
Debug script to investigate space loss in bold text runs during translation.

This script demonstrates the issue described as:
- Example: "ciao **sono** io" becomes "hello**This is**me" (spaces around bold text are lost)
"""

def debug_space_issue():
    """Debug the space loss issue in text runs"""
    
    # Simulate what happens in the translator
    print("=== SPACE LOSS ISSUE DEBUG ===\n")
    
    # Simulate a paragraph with mixed formatting: "ciao **sono** io"
    # In PowerPoint, this would be represented as 3 runs:
    # Run 1: "ciao " (normal)
    # Run 2: "sono" (bold)  
    # Run 3: " io" (normal)
    
    runs = [
        {"text": "ciao ", "bold": False},
        {"text": "sono", "bold": True},
        {"text": " io", "bold": False}
    ]
    
    print("Original runs:")
    for i, run in enumerate(runs):
        print(f"  Run {i+1}: '{run['text']}' (Bold: {run['bold']})")
    
    # Show what the full text looks like
    full_text = "".join(run["text"] for run in runs)
    print(f"\nFull original text: '{full_text}'")
    print(f"Visual representation: ciao **sono** io")
    
    # Now simulate the translation process
    print("\n=== TRANSLATION PROCESS ===")
    
    def mock_translate(text):
        """Mock translation function that shows the issue"""
        print(f"  Translating: '{text}' -> ", end="")
        
        # Simple Italian to English translations
        translations = {
            "ciao ": "hello ",
            "sono": "This is", 
            " io": " me"
        }
        
        result = translations.get(text, text)
        print(f"'{result}'")
        return result
    
    # Translate each run individually (current approach)
    print("\nTranslating each run individually:")
    translated_runs = []
    for i, run in enumerate(runs):
        translated_text = mock_translate(run["text"])
        translated_runs.append({
            "text": translated_text,
            "bold": run["bold"]
        })
    
    # Show the result
    print("\nTranslated runs:")
    for i, run in enumerate(translated_runs):
        print(f"  Run {i+1}: '{run['text']}' (Bold: {run['bold']})")
    
    # Show what the full text looks like after translation
    full_translated = "".join(run["text"] for run in translated_runs)
    print(f"\nFull translated text: '{full_translated}'")
    print(f"Visual representation: hello **This is** me")
    
    # Show the problem
    print("\n=== PROBLEM ANALYSIS ===")
    print("✅ EXPECTED: 'hello **This is** me' (spaces preserved)")
    print("✅ ACTUAL:   'hello **This is** me' (spaces preserved)")
    print("✅ In this case, the translation works correctly!")
    
    # Now let's test a case where the AI might return text without proper spacing
    print("\n=== TESTING PROBLEMATIC AI RESPONSE ===")
    
    def problematic_translate(text):
        """Mock AI translation that doesn't preserve spaces properly"""
        print(f"  Translating: '{text}' -> ", end="")
        
        # AI sometimes returns text without proper spacing
        translations = {
            "ciao ": "hello",  # Missing trailing space!
            "sono": "This is", 
            " io": "me"        # Missing leading space!
        }
        
        result = translations.get(text, text)
        print(f"'{result}'")
        return result
    
    print("\nTranslating with problematic AI (spaces lost):")
    problematic_runs = []
    for i, run in enumerate(runs):
        translated_text = problematic_translate(run["text"])
        problematic_runs.append({
            "text": translated_text,
            "bold": run["bold"]
        })
    
    # Show the problematic result
    print("\nProblematic translated runs:")
    for i, run in enumerate(problematic_runs):
        print(f"  Run {i+1}: '{run['text']}' (Bold: {run['bold']})")
    
    # Show what the full text looks like
    full_problematic = "".join(run["text"] for run in problematic_runs)
    print(f"\nFull problematic text: '{full_problematic}'")
    print(f"Visual representation: hello**This is**me")
    
    print("\n=== PROBLEM IDENTIFIED ===")
    print("❌ EXPECTED: 'hello **This is** me' (spaces around bold)")
    print("❌ ACTUAL:   'hello**This is**me' (no spaces around bold)")
    print("❌ ISSUE: AI translation is not preserving leading/trailing spaces in runs!")
    
    # Show potential solutions
    print("\n=== POTENTIAL SOLUTIONS ===")
    print("1. Pre-process: Detect and preserve leading/trailing spaces")
    print("2. Post-process: Restore original spacing patterns")
    print("3. Context-aware: Send full paragraph but map back to runs")
    print("4. AI instruction: Explicitly tell AI to preserve all spacing")

if __name__ == "__main__":
    debug_space_issue()
