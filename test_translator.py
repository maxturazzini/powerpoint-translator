#!/usr/bin/env python3
import sys
import os
sys.path.append('.')
from translate_powerpoint import PowerPointTranslator
import logging

# Set up logging to see what happens during translation
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def mock_translate(text):
    """Mock translation function for testing - just adds '[TRANSLATED]' prefix"""
    if not text or not text.strip():
        return text
    return f"[TRANSLATED] {text}"

def test_translator():
    """Test the translator with the sample file"""
    try:
        # Initialize translator with dummy API key
        translator = PowerPointTranslator(
            api_key="dummy_key_for_testing",
            model="gpt-4o-mini",
            translate_notes=True,
            skip_hidden_slides=False
        )
        
        input_path = 'sample_pptx/renewable_energy_sample_translation.pptx'
        output_path = 'sample_pptx/test_output_formatting.pptx'
        
        print(f"🧪 Testing translator with: {input_path}")
        print(f"📤 Output will be saved to: {output_path}")
        
        # Use mock translation function instead of OpenAI
        translator.translate_presentation(input_path, output_path, mock_translate)
        
        print("✅ Translation completed successfully!")
        print(f"📁 Check the output file: {output_path}")
        
        # Now analyze the output to see formatting preservation
        print("\n🔍 Analyzing output formatting...")
        from analyze_sample import analyze_presentation
        print("\n" + "="*60)
        print("ANALYZING OUTPUT FILE:")
        print("="*60)
        analyze_presentation(output_path)
        
    except Exception as e:
        print(f"❌ Error during translation test: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_translator()