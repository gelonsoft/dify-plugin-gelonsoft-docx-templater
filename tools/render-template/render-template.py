import io
import json
from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from docxtpl import DocxTemplate


class RenderTemplateTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        """
        Generate a .docx file from a template and JSON context data.
        """
        # Extract parameters
        template_file = tool_parameters.get("template")
        context_str = tool_parameters.get("context", "{}")
        filename = tool_parameters.get("filename", "output")

        # Parse JSON context
        try:
            context = json.loads(context_str)
        except json.JSONDecodeError as e:
            yield self.create_text_message(f"Error: Invalid JSON context - {str(e)}")
            return

        # Validate template file
        template_path = template_file.path if hasattr(template_file, "path") else None
        if not template_path:
            yield self.create_text_message("Error: No template file provided")
            return

        # Render the DOCX template
        try:
            doc = DocxTemplate(template_path)
            doc.render(context)

            output_buffer = io.BytesIO()
            doc.save(output_buffer)
            output_buffer.seek(0)
            file_bytes = output_buffer.getvalue()
        except Exception as e:
            yield self.create_text_message(f"Error rendering template: {str(e)}")
            return

        # Ensure filename has .docx extension
        if not filename.endswith(".docx"):
            filename += ".docx"

        # Yield the generated document as a blob
        yield self.create_blob_message(
            blob=file_bytes,
            meta={
                "mime_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "filename": filename,
            },
        )

        # Yield a text confirmation
        yield self.create_text_message(f"Document '{filename}' generated successfully.")
