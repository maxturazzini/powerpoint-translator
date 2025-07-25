from typing import Optional, List, Dict, Tuple
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from docx.table import Table, _Cell
from word_formatting.manager import WordFormattingManager
from .text_processor import WordTextProcessor
import logging

logger = logging.getLogger('word_translator')

class EnhancedDocumentProcessor:
    """Enhanced document processor with run-preserving translation"""
    
    def __init__(self, formatting_manager: WordFormattingManager):
        self.formatting_manager = formatting_manager
        self.text_processor = WordTextProcessor()
        
    def process_paragraph(self, paragraph: Paragraph, translate_func) -> None:
        """
        Process a paragraph while preserving formatting at run level.
        This is the core function that maintains complex formatting.
        """
        try:
            logger.info(f"Processing paragraph with {len(paragraph.runs)} runs")
            
            # Check if paragraph has content to translate
            if not paragraph.text.strip():
                logger.info("Skipping empty paragraph")
                return
            
            # Use enhanced run-level translation
            self._translate_paragraph_runs(paragraph, translate_func)
            
        except Exception as e:
            logger.error(f"Error processing paragraph: {str(e)}")
            raise

    def _translate_paragraph_runs(self, paragraph: Paragraph, translate_func) -> None:
        """
        Translate each run in a paragraph individually, preserving formatting boundaries.
        This is the core innovation adapted from the PowerPoint translator.
        """
        if not paragraph.runs:
            return
            
        logger.info(f"Translating paragraph with {len(paragraph.runs)} runs")
        
        # Strategy: Translate each run individually to preserve formatting boundaries
        for run_idx, run in enumerate(paragraph.runs):
            if run.text and run.text.strip():
                original_text = run.text
                
                # Validate text before translation
                if not self.formatting_manager._should_translate_text(original_text):
                    logger.info(f"  Run {run_idx + 1}: Skipping '{original_text}' (validation failed)")
                    continue
                
                # Preserve whitespace patterns for proper spacing
                leading_spaces = original_text[:len(original_text) - len(original_text.lstrip())]
                trailing_spaces = original_text[len(original_text.rstrip()):]
                core_text = original_text.strip()
                
                # Translate only the core text content
                translated_core = translate_func(core_text)
                
                # Reconstruct with original spacing
                translated_text = leading_spaces + translated_core + trailing_spaces
                
                # Update the run's text while preserving all formatting
                run.text = translated_text
                
                logger.info(f"  Run {run_idx + 1}: '{original_text}' -> '{translated_text}'")
        
        logger.info(f"Paragraph translation completed with preserved formatting")

    def process_table_cell(self, cell: _Cell, translate_func) -> None:
        """Process all paragraphs in a table cell"""
        try:
            logger.info("Processing table cell")
            
            for para_idx, paragraph in enumerate(cell.paragraphs):
                if paragraph.text.strip():
                    logger.info(f"Processing cell paragraph {para_idx + 1}")
                    self.process_paragraph(paragraph, translate_func)
                    
        except Exception as e:
            logger.error(f"Error processing table cell: {str(e)}")
            raise

    def process_table(self, table: Table, translate_func) -> None:
        """Process all cells in a table"""
        try:
            logger.info("Processing table")
            
            for row_idx, row in enumerate(table.rows):
                for cell_idx, cell in enumerate(row.cells):
                    logger.info(f"Processing table cell [{row_idx + 1},{cell_idx + 1}]")
                    self.process_table_cell(cell, translate_func)
                    
        except Exception as e:
            logger.error(f"Error processing table: {str(e)}")
            raise

    def _context_aware_translation(self, paragraph: Paragraph, translate_func) -> None:
        """
        Advanced strategy: Translate with context awareness.
        Sends full paragraph to translator but maps result back to runs.
        This can be used for more sophisticated translation approaches.
        """
        if not paragraph.runs:
            return
            
        # Build full paragraph text with run boundary information
        full_text = ""
        run_boundaries = []
        current_pos = 0
        
        for run_idx, run in enumerate(paragraph.runs):
            run_text = run.text
            run_boundaries.append({
                'start': current_pos,
                'end': current_pos + len(run_text),
                'text': run_text,
                'run': run,
                'formatting': self.formatting_manager.collect_run_formatting(run, paragraph)
            })
            full_text += run_text
            current_pos += len(run_text)
        
        if not full_text.strip():
            return
            
        # For now, fall back to individual run translation
        # This could be enhanced to do intelligent text redistribution
        for boundary in run_boundaries:
            if boundary['text'].strip():
                if self.formatting_manager._should_translate_text(boundary['text']):
                    boundary['run'].text = translate_func(boundary['text'])

    def _intelligent_run_redistribution(self, paragraph: Paragraph, translated_text: str) -> None:
        """
        Intelligent redistribution of translated text back to runs.
        This handles cases where word boundaries might change during translation.
        Currently a placeholder for future enhancement.
        """
        # This is a complex feature that could be implemented later
        # For now, the run-by-run approach works well for most cases
        pass

    def _preserve_paragraph_structure(self, paragraph: Paragraph) -> Dict:
        """Store paragraph-level structure and formatting"""
        try:
            return {
                'alignment': paragraph.alignment,
                'style': paragraph.style.name if paragraph.style else None,
                'run_count': len(paragraph.runs),
                'space_before': paragraph.paragraph_format.space_before,
                'space_after': paragraph.paragraph_format.space_after,
                'line_spacing': paragraph.paragraph_format.line_spacing,
                'left_indent': paragraph.paragraph_format.left_indent,
                'right_indent': paragraph.paragraph_format.right_indent,
                'first_line_indent': paragraph.paragraph_format.first_line_indent
            }
        except Exception as e:
            logger.warning(f"Error preserving paragraph structure: {str(e)}")
            return {}

    def _restore_paragraph_structure(self, paragraph: Paragraph, structure: Dict) -> None:
        """Restore paragraph-level structure and formatting"""
        try:
            if structure.get('alignment') is not None:
                paragraph.alignment = structure['alignment']
            
            if structure.get('style'):
                try:
                    paragraph.style = structure['style']
                except Exception:
                    logger.warning(f"Could not restore paragraph style: {structure['style']}")
            
            para_format = paragraph.paragraph_format
            if structure.get('space_before') is not None:
                para_format.space_before = structure['space_before']
            if structure.get('space_after') is not None:
                para_format.space_after = structure['space_after']
            if structure.get('line_spacing') is not None:
                para_format.line_spacing = structure['line_spacing']
            if structure.get('left_indent') is not None:
                para_format.left_indent = structure['left_indent']
            if structure.get('right_indent') is not None:
                para_format.right_indent = structure['right_indent']
            if structure.get('first_line_indent') is not None:
                para_format.first_line_indent = structure['first_line_indent']
                
        except Exception as e:
            logger.warning(f"Error restoring paragraph structure: {str(e)}")

    def process_header_footer(self, header_footer, translate_func) -> None:
        """Process paragraphs in headers or footers"""
        try:
            logger.info("Processing header/footer")
            
            for para_idx, paragraph in enumerate(header_footer.paragraphs):
                if paragraph.text.strip():
                    logger.info(f"Processing header/footer paragraph {para_idx + 1}")
                    self.process_paragraph(paragraph, translate_func)
                    
        except Exception as e:
            logger.error(f"Error processing header/footer: {str(e)}")
            raise

    def validate_translation_quality(self, original_paragraph: Paragraph, 
                                   translated_paragraph: Paragraph) -> List[str]:
        """Validate that translation preserved formatting structure"""
        warnings = []
        
        try:
            original_runs = len(original_paragraph.runs)
            translated_runs = len(translated_paragraph.runs)
            
            if original_runs != translated_runs:
                warnings.append(f"Run count mismatch: {original_runs} -> {translated_runs}")
            
            # Check if basic formatting is preserved
            for i, (orig_run, trans_run) in enumerate(
                zip(original_paragraph.runs, translated_paragraph.runs)
            ):
                if (orig_run.font.bold != trans_run.font.bold or
                    orig_run.font.italic != trans_run.font.italic):
                    warnings.append(f"Formatting mismatch in run {i + 1}")
            
        except Exception as e:
            warnings.append(f"Validation error: {str(e)}")
        
        return warnings