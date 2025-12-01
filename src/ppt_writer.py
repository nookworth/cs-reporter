"""
PowerPoint generator using python-pptx.
Replaces placeholders in a template with data from Excel.
"""

from pathlib import Path
from datetime import datetime
from pptx import Presentation
from pptx.util import Pt
from copy import deepcopy
import re


class PowerPointWriter:
    """Generates PowerPoint reports by replacing placeholders in a template."""

    def __init__(self, config):
        """
        Initialize the PowerPoint writer.

        Args:
            config: Configuration dictionary with template path
        """
        self.config = config
        self.template_path = Path(config['template_path'])
        self.output_dir = Path(config.get('output_dir', 'output'))

        # Ensure output directory exists
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def generate_report(self, data):
        """
        Generate a PowerPoint report by replacing placeholders with data.

        Args:
            data: Dictionary mapping field names to values

        Returns:
            Path to the generated PowerPoint file
        """
        # Load the template
        prs = Presentation(self.template_path)

        # Replace placeholders in all slides
        replacements_made = self._replace_placeholders(prs, data)

        print(f"  Made {replacements_made} replacements")

        # Generate output filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"report_{timestamp}.pptx"
        output_path = self.output_dir / output_filename

        # Save the presentation
        prs.save(str(output_path))

        return output_path

    def _replace_placeholders(self, presentation, data):
        """
        Replace all {{placeholder}} tokens in the presentation with data values.

        Args:
            presentation: python-pptx Presentation object
            data: Dictionary of field names to values (includes tables as lists)

        Returns:
            Number of replacements made
        """
        replacements = 0

        # Separate table data from scalar data
        table_data = {k: v for k, v in data.items() if isinstance(v, list)}
        scalar_data = {k: v for k, v in data.items() if not isinstance(v, list)}

        # Iterate through all slides
        for slide in presentation.slides:
            # Iterate through all shapes in the slide
            for shape in slide.shapes:
                if hasattr(shape, "text_frame"):
                    replacements += self._replace_in_text_frame(
                        shape.text_frame, scalar_data
                    )

                # Check tables
                if shape.has_table:
                    # Check if this table has a dynamic table placeholder
                    table_replacement = self._populate_table(shape.table, table_data)
                    if table_replacement:
                        replacements += table_replacement
                    else:
                        # Regular cell replacement for static tables
                        for row in shape.table.rows:
                            for cell in row.cells:
                                replacements += self._replace_in_text_frame(
                                    cell.text_frame, scalar_data
                                )

        return replacements

    def _populate_table(self, table, table_data):
        """
        Populate a table with dynamic rows if it contains a {{table:name}} placeholder.

        Args:
            table: python-pptx Table object
            table_data: Dictionary of table names to lists of row data

        Returns:
            Number of replacements made, or None if not a dynamic table
        """
        # Check if the first cell of the first row contains {{table:table_name}}
        first_cell_text = table.rows[0].cells[0].text
        table_match = re.search(r'\{\{table:(\w+)\}\}', first_cell_text)

        if not table_match:
            return None

        table_name = table_match.group(1)

        if table_name not in table_data:
            print(f"  Warning: Table '{table_name}' not found in data")
            return 0

        rows_data = table_data[table_name]

        if not rows_data:
            print(f"  Warning: Table '{table_name}' has no data rows")
            return 0

        # The first row is the template
        template_row_idx = 0
        template_row = table.rows[template_row_idx]

        # Extract template cell properties and placeholders
        template_cells = []
        for cell in template_row.cells:
            cell_info = {
                'text': cell.text,
                'text_frame': cell.text_frame
            }
            template_cells.append(cell_info)

        # Delete all existing rows
        # We'll rebuild the table from scratch
        # Note: python-pptx doesn't support deleting rows easily,
        # so we'll clear and repopulate existing rows

        # Determine how many rows we need
        num_data_rows = len(rows_data)
        num_existing_rows = len(table.rows)

        # Add more rows if needed
        while len(table.rows) < num_data_rows:
            table.add_row()

        # Remove extra rows if needed (clear them at least)
        # python-pptx doesn't easily support row deletion, so we keep them

        # Populate each row with data
        replacements = 0
        for row_idx, row_data in enumerate(rows_data):
            row = table.rows[row_idx]

            for col_idx, cell in enumerate(row.cells):
                # Get template text for this column
                template_text = template_cells[col_idx]['text']

                # Replace placeholders in the template text
                new_text = template_text
                for field_name, value in row_data.items():
                    placeholder = f"{{{{{field_name}}}}}"
                    if placeholder in new_text:
                        new_text = new_text.replace(placeholder, str(value))
                        replacements += 1

                # Also remove the {{table:name}} marker from first cell
                new_text = re.sub(r'\{\{table:\w+\}\}', '', new_text).strip()

                # Set the cell text
                cell.text = new_text

        print(f"  Populated table '{table_name}' with {num_data_rows} rows")

        return replacements

    def _replace_in_text_frame(self, text_frame, data):
        """
        Replace placeholders in a text frame.

        Args:
            text_frame: python-pptx TextFrame object
            data: Dictionary of field names to values

        Returns:
            Number of replacements made
        """
        replacements = 0

        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                original_text = run.text

                # Replace all {{field_name}} patterns
                for field_name, value in data.items():
                    placeholder = f"{{{{{field_name}}}}}"
                    if placeholder in run.text:
                        run.text = run.text.replace(placeholder, str(value))
                        replacements += 1

        return replacements
