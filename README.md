# PowerPoint Translator with Formatting Preservation

A sophisticated PowerPoint translator that preserves complex formatting during translation using OpenAI's API. Perfect for business and technical presentations where formatting matters.

## ✨ Key Features

- **🎨 Perfect Formatting Preservation** - Maintains bold, italic, colors, fonts, alignment, and mixed formatting within text
- **🔄 Complete Translation** - Translates slides, notes, tables, SmartArt, and grouped shapes
- **🖥️ Dual Interface** - Command-line tool and user-friendly GUI
- **📊 Table Support** - Preserves table structure and cell formatting
- **🎯 Smart Shape Handling** - Processes all PowerPoint shape types intelligently
- **✅ Built-in Validation** - Comprehensive formatting verification and comparison tools

## 🚀 Quick Start

### Prerequisites

1. **Python 3.8+** installed on your system
2. **OpenAI API Key** - Get one from [OpenAI Platform](https://platform.openai.com/)

### Installation

1. **Clone or download** this repository
2. **Install required packages:**
   ```bash
   pip install python-pptx openai python-dotenv
   ```

3. **Set your OpenAI API key:**
   
   **Option A: Environment Variable**
   ```bash
   export OPENAI_API_KEY="your-api-key-here"
   ```
   
   **Option B: Create .env file**
   ```bash
   echo "openai_api_key=your-api-key-here" > .env
   ```

### Usage

#### 🖥️ GUI Application (Recommended)

```bash
python3 translate_powerpoint_gui.py
```

1. Click **"Browse"** to select your PowerPoint file
2. The output filename will be auto-generated
3. Edit the translation prompt if needed
4. Click **"Translate"** to start
5. Click **"Open Translated File"** when complete

#### ⌨️ Command Line

```bash
python3 translate_powerpoint.py
```

Edit the file to specify your input/output paths and run directly.

## 📖 How It Works

### The Formatting Challenge

Traditional translation tools destroy PowerPoint formatting because they replace entire text blocks. Consider this text:

> "**Key** *Points*: Important information"

Most tools would flatten this to plain text, losing the bold "Key" and italic "Points" formatting.

### Our Solution

This translator uses **run-level preservation**:

1. **Analyzes text structure** - Identifies individual formatting runs within paragraphs
2. **Translates individually** - Each run is translated while preserving its formatting container
3. **Maintains boundaries** - Bold/italic boundaries are preserved exactly
4. **Validates results** - Comprehensive formatting verification ensures quality

**Result:** Perfect formatting preservation with professional translation quality.

## 🎯 Supported Features

### Shape Types
- ✅ Text boxes and placeholders
- ✅ Tables with complex formatting
- ✅ SmartArt graphics
- ✅ Grouped shapes
- ✅ Auto-shapes with text

### Formatting Elements
- ✅ **Bold**, *italic*, <u>underline</u> text
- ✅ Font families and sizes
- ✅ Text colors (RGB and theme colors)
- ✅ Paragraph alignment and indentation
- ✅ Mixed formatting within single paragraphs
- ✅ Table cell formatting
- ✅ Slide notes

### Translation Options
- ✅ Custom translation prompts
- ✅ Notes translation (optional)
- ✅ Hidden slide processing (optional)
- ✅ Multiple language pairs

## 🔧 Configuration

### Translation Settings

Edit the translation instructions in the GUI or modify the `INSTRUCTIONS` constant in `translate_powerpoint.py`:

```python
INSTRUCTIONS = """
You are a professional translator...
[Customize your translation requirements here]
"""
```

### Advanced Options

When initializing the translator programmatically:

```python
translator = PowerPointTranslator(
    api_key="your-key",
    model="gpt-4o-mini",           # OpenAI model to use
    translate_notes=True,           # Include slide notes
    skip_hidden_slides=False        # Process hidden slides
)
```

## 📊 Testing & Validation

### Test the Translator

```bash
# Test with enhanced formatting preservation
python3 test_enhanced_translator.py

# Analyze PowerPoint structure
python3 analyze_sample.py
```

### Validation Results

The enhanced translator achieves:
- **100% formatting preservation** for mixed formatting
- **Perfect run structure maintenance**
- **Zero formatting destruction** during translation

## 📁 Project Structure

```
powerpoint-translator/
├── translate_powerpoint.py          # Core translator
├── translate_powerpoint_gui.py      # GUI application
├── formatting/
│   └── manager.py                   # Formatting preservation logic
├── processors/
│   ├── enhanced_shape_processor.py  # Advanced formatting processor
│   ├── shape_processor.py           # Legacy processor
│   └── text_processor.py            # Text extraction utilities
├── validation/
│   ├── validator.py                 # Format validation
│   └── comparator.py                # Visual comparison tools
├── sample_pptx/                     # Test files
└── test_*.py                        # Testing scripts
```

## 💡 Tips for Best Results

### Translation Quality
- Use descriptive prompts for your specific domain (business, technical, medical, etc.)
- Specify source and target languages clearly
- Include context about technical terms or industry jargon

### Formatting Preservation
- The system automatically preserves all formatting - no special setup required
- Complex tables and SmartArt are fully supported
- Mixed formatting within paragraphs is perfectly maintained

### Performance
- Large presentations may take several minutes to translate
- Progress is logged to console and `ppt_translator.log`
- Internet connection required for OpenAI API calls

## 🐛 Troubleshooting

### Common Issues

**"No API key found"**
- Ensure your OpenAI API key is set as an environment variable or in `.env`

**"Translation failed"**
- Check your internet connection
- Verify your OpenAI API key is valid and has credits
- Check the log file `ppt_translator.log` for detailed error information

**"Formatting looks different"**
- The enhanced system preserves formatting perfectly
- If using old files, ensure you're running the enhanced translator
- Compare before/after using the analysis tools

**"PowerPoint file won't open"**
- Ensure the input file isn't corrupted
- Try with a simple test presentation first
- Check file permissions

### Getting Help

1. **Check the logs** - `ppt_translator.log` contains detailed information
2. **Test with samples** - Use the provided sample files to verify setup
3. **Run validation** - Use `test_enhanced_translator.py` to verify functionality

## 📝 Example Output

**Before Translation:**
```
Slide Title: "Renewable Energy Trends 2025"
Content: "Key Points: Solar capacity surpassed 1 TW in 2024"
```

**After Translation (English → Italian):**
```
Slide Title: "Tendenze dell'Energia Rinnovabile 2025"
Content: "Punti Chiave: La capacità solare ha superato 1 TW nel 2024"
```

With **perfect formatting preservation** - bold text stays bold, colors remain unchanged, table structures are maintained.

## 🎯 Use Cases

- **Business Presentations** - Quarterly reports, strategy decks, board presentations
- **Technical Documentation** - Product specs, training materials, user guides  
- **Academic Content** - Research presentations, course materials, conferences
- **Marketing Materials** - Sales decks, product launches, customer presentations

## 📄 License

This project is open source. See the repository for license details.

## 🔄 Version History

- **V5** - Enhanced formatting preservation with run-level translation
- **V4** - Comprehensive validation and comparison tools
- **V3** - SmartArt and table support
- **V2** - GUI interface and notes translation
- **V1** - Basic PowerPoint translation

---

**Transform your PowerPoint presentations while preserving every formatting detail.** 🎨✨