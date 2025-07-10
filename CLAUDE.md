# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a PowerPoint translator that preserves complex formatting during translation using OpenAI's API. The system translates PowerPoint presentations while maintaining exact formatting properties including mixed bold/italic text within single paragraphs, colors, fonts, alignment, and table formatting.

## Development Commands

### Running the Translator

**Command Line (Core Translator):**
```bash
python3 translate_powerpoint.py
```

**GUI Application:**
```bash
python3 translate_powerpoint_gui.py
```

### Testing and Validation

**Run Enhanced Translator Test:**
```bash
python3 test_enhanced_translator.py
```

**Run Original Translator Test:**
```bash
python3 test_translator.py
```

**Analyze Sample PowerPoint Formatting:**
```bash
python3 analyze_sample.py
```

### Environment Setup

Set your OpenAI API key as an environment variable:
```bash
export OPENAI_API_KEY="your-api-key-here"
```

Or create a `.env` file in the parent directory with:
```
openai_api_key=your-api-key-here
```

## Architecture Overview

### Core Translation System

The translator uses a **dual-processor architecture** with enhanced formatting preservation:

1. **Original System** (`processors/shape_processor.py`) - Legacy processor that uses destructive text clearing
2. **Enhanced System** (`processors/enhanced_shape_processor.py`) - New run-preserving processor that maintains formatting

**Key Architectural Decision:** The enhanced system translates individual text runs in-place without clearing text frames, preserving the original paragraph/run structure that contains formatting information.

### Component Hierarchy

```
PowerPointTranslator (main orchestrator)
├── FormattingManager (formatting/manager.py)
│   ├── TextRunFormatting (dataclass for run-level formatting)
│   └── Format capture/restore logic
├── EnhancedShapeProcessor (processors/enhanced_shape_processor.py) [CURRENT]
│   ├── Run-by-run translation (preserves structure)
│   ├── Table cell processing
│   └── SmartArt handling
├── ShapeProcessor (processors/shape_processor.py) [LEGACY]
│   └── Original destructive approach
├── TextProcessor (processors/text_processor.py)
│   └── Content extraction utilities
└── Validation (validation/)
    ├── FormatValidator - Structure validation
    └── VisualComparator - Formatting comparison
```

### Critical Formatting Preservation Logic

The **EnhancedShapeProcessor** implements the core innovation:

- **`_translate_paragraph_runs()`** - Translates each text run individually while preserving formatting containers
- **No `text_frame.clear()`** - Avoids destructive clearing that destroys run structure
- **In-place text updates** - Updates `run.text` directly maintaining formatting boundaries

This approach solves the fundamental issue where mixed formatting (e.g., "**Bold** and *Italic*") was being flattened into single runs.

### Translation Flow

1. **Shape Detection** - Identifies text frames, tables, SmartArt, groups
2. **Run-Level Processing** - For each paragraph, processes individual runs
3. **Format Preservation** - Maintains original run boundaries during translation
4. **Validation** - Compares formatting before/after translation

### Shape Type Handling

- **Text Frames/Placeholders** - Run-by-run translation preserving formatting
- **Tables** - Cell-by-cell processing without extra line breaks
- **SmartArt** - Enhanced text element handling
- **Groups** - Recursive processing of nested shapes

## File Structure Context

- **`translate_powerpoint.py`** - Main translator class and CLI interface
- **`translate_powerpoint_gui.py`** - Tkinter GUI wrapper
- **`formatting/manager.py`** - Core formatting capture/restore logic
- **`processors/enhanced_shape_processor.py`** - Enhanced run-preserving processor (USE THIS)
- **`processors/shape_processor.py`** - Legacy processor (formatting issues)
- **`sample_pptx/`** - Test files for validation
- **`test_enhanced_translator.py`** - Validation script demonstrating 100% formatting preservation

## Translation Instructions

The system includes predefined translation instructions for English→Italian business/technical presentations. These can be customized in the `INSTRUCTIONS` constant or via the GUI's system prompt editor.

## Testing Strategy

The project includes comprehensive formatting preservation tests:

- **Sample Analysis** - `analyze_sample.py` examines PowerPoint structure
- **Comparison Testing** - Before/after formatting validation
- **Mixed Formatting Tests** - Validates preservation of complex text formatting
- **Success Metrics** - Measures run count preservation and formatting accuracy

## Important Notes

- **Always use EnhancedShapeProcessor** for new development - it provides 100% formatting preservation
- **API Key Security** - Never commit API keys; use environment variables or .env files
- **Logging** - All operations log to both console and `ppt_translator.log`
- **Mock Translation** - Use mock functions for testing without API calls
- **Sample Files** - Test against `sample_pptx/renewable_energy_sample_translation.pptx` for validation

## Known Issues and Improvements

- **Text Translation Validation**: ✅ **COMPLETED**
  - ~~Sometimes no text is passed and the AI returns an error response~~
  - ~~Need to implement validation to avoid translating:~~
    - ~~Empty strings~~
    - ~~Full uppercase strings (Acronyms or single letters)~~
  - **Implementation**: Added `_should_translate_text()` method in `EnhancedShapeProcessor` that validates text before translation
  - **Features**: Skips empty strings, whitespace, single characters, uppercase acronyms, and pure numbers
  - **Location**: `processors/enhanced_shape_processor.py:259-292`

## Translation Behavior Updates

- Do not translate links. App now responds: "I'm sorry, but I cannot access external links or content. Please provide the text you would like me to translate, and I'll be happy to assist you."

## Future Development Ideas

- **Prompt Optimization**:
  - **Todo 3**: Propose a way to compact the system input before the user clicks to launch
    - Maybe add a 'compact prompt' button in GUI that runs an analysis with GPT 4.1 or o4 and minimize token usage
    - **Status**: Pending implementation

## Known Formatting Issues

- **Todo 5**: Detected issue with bold text formatting: ✅ **COMPLETED**
  - ~~Sometimes in bold words inside a run, spaces get lost~~
  - ~~Example: "ciao **sono** io" becomes "hello**This is**me"~~
  - ~~Requires investigation of run-level text preservation mechanism~~
  - **Root Cause**: AI translation service doesn't preserve leading/trailing spaces in individual text runs
  - **Solution**: Enhanced `_translate_paragraph_runs()` method to preserve whitespace patterns
  - **Implementation**: Extract and preserve leading/trailing spaces, translate only core text, then reconstruct
  - **Location**: `processors/enhanced_shape_processor.py:122-131`

- **Todo 6**: Format is revert to standard fonts when no font is specified. Evaluate if this is a bug or a pptx file problem
  - **Status**: Low priority, requires investigation
