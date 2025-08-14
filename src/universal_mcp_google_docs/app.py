from typing import Any

from universal_mcp.applications.application import APIApplication
from universal_mcp.integrations import Integration


class GoogleDocsApp(APIApplication):
    def __init__(self, integration: Integration) -> None:
        super().__init__(name="google-docs", integration=integration)
        self.base_api_url = "https://docs.googleapis.com/v1/documents"

    def create_document(self, title: str) -> dict[str, Any]:
        """
        Creates a new blank Google Document with the specified title and returns the API response.

        Args:
            title: The title for the new Google Document to be created

        Returns:
            A dictionary containing the Google Docs API response with document details and metadata

        Raises:
            HTTPError: If the API request fails due to network issues, authentication errors, or invalid parameters
            RequestException: If there are connection errors or timeout issues during the API request

        Tags:
            create, document, api, important, google-docs, http
        """
        url = self.base_api_url
        document_data = {"title": title}
        response = self._post(url, data=document_data)
        response.raise_for_status()
        return response.json()

    def get_document(self, document_id: str) -> dict[str, Any]:
        """
        Retrieves the latest version of a specified document from the Google Docs API.

        Args:
            document_id: The unique identifier of the document to retrieve

        Returns:
            A dictionary containing the document data from the Google Docs API response

        Raises:
            HTTPError: If the API request fails or the document is not found
            JSONDecodeError: If the API response cannot be parsed as JSON

        Tags:
            retrieve, read, api, document, google-docs, important
        """
        url = f"{self.base_api_url}/{document_id}"
        response = self._get(url)
        return response.json()

    def add_content(
        self, document_id: str, content: str, index: int = 1
    ) -> dict[str, Any]:
        """
        Adds text content at a specified position in an existing Google Document via the Google Docs API.

        Args:
            document_id: The unique identifier of the Google Document to be updated
            content: The text content to be inserted into the document
            index: The zero-based position in the document where the text should be inserted (default: 1)

        Returns:
            A dictionary containing the Google Docs API response after performing the batch update operation

        Raises:
            HTTPError: When the API request fails, such as invalid document_id or insufficient permissions
            RequestException: When there are network connectivity issues or API endpoint problems

        Tags:
            update, insert, document, api, google-docs, batch, content-management, important
        """
        url = f"{self.base_api_url}/{document_id}:batchUpdate"
        batch_update_data = {
            "requests": [
                {"insertText": {"location": {"index": index}, "text": content}}
            ]
        }
        response = self._post(url, data=batch_update_data)
        response.raise_for_status()
        return response.json()

    def style_text(
        self,
        document_id: str,
        start_index: int,
        end_index: int,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        strikethrough: bool | None = None,
        small_caps: bool | None = None,
        font_size: float | None = None,
        font_family: str | None = None,
        font_weight: int | None = None,
        foreground_color: dict[str, float] | None = None,
        background_color: dict[str, float] | None = None,
        link_url: str | None = None,
    ) -> dict[str, Any]:
        """
        Applies text styling to a specified range of text in a Google Document.

        Args:
            document_id: The unique identifier of the Google Document to be updated
            start_index: The zero-based start index of the text range to style
            end_index: The zero-based end index of the text range to style (exclusive)
            bold: Whether the text should be bold
            italic: Whether the text should be italicized
            underline: Whether the text should be underlined
            strikethrough: Whether the text should be struck through
            small_caps: Whether the text should be in small capital letters
            font_size: The font size in points (e.g., 12.0 for 12pt)
            font_family: The font family name (e.g., "Arial", "Times New Roman")
            font_weight: The font weight (e.g., 400 for normal, 700 for bold)
            foreground_color: RGB color dict with 'red', 'green', 'blue' values (0.0 to 1.0)
            background_color: RGB color dict with 'red', 'green', 'blue' values (0.0 to 1.0)
            link_url: URL to make the text a hyperlink

        Returns:
            A dictionary containing the Google Docs API response after performing the batch update operation

        Raises:
            HTTPError: When the API request fails, such as invalid document_id or insufficient permissions
            RequestException: When there are network connectivity issues or API endpoint problems

        Tags:
            style, format, text, document, api, google-docs, batch, formatting, important
        """
        url = f"{self.base_api_url}/{document_id}:batchUpdate"
        
        # Build the text style object
        text_style = {}
        
        if bold is not None:
            text_style["bold"] = bold
        if italic is not None:
            text_style["italic"] = italic
        if underline is not None:
            text_style["underline"] = underline
        if strikethrough is not None:
            text_style["strikethrough"] = strikethrough
        if small_caps is not None:
            text_style["smallCaps"] = small_caps
        if font_size is not None:
            text_style["fontSize"] = {"magnitude": font_size, "unit": "PT"}
        if font_family is not None:
            weighted_font_family = {"fontFamily": font_family}
            if font_weight is not None:
                weighted_font_family["weight"] = str(font_weight)
            text_style["weightedFontFamily"] = weighted_font_family
        if foreground_color is not None:
            text_style["foregroundColor"] = {
                "color": {
                    "rgbColor": {
                        "red": foreground_color.get("red", 0.0),
                        "green": foreground_color.get("green", 0.0),
                        "blue": foreground_color.get("blue", 0.0)
                    }
                }
            }
        if background_color is not None:
            text_style["backgroundColor"] = {
                "color": {
                    "rgbColor": {
                        "red": background_color.get("red", 0.0),
                        "green": background_color.get("green", 0.0),
                        "blue": background_color.get("blue", 0.0)
                    }
                }
            }
        if link_url is not None:
            text_style["link"] = {"url": link_url}
        
        batch_update_data = {
            "requests": [
                {
                    "updateTextStyle": {
                        "range": {
                            "startIndex": start_index,
                            "endIndex": end_index
                        },
                        "textStyle": text_style,
                        "fields": "*"
                    }
                }
            ]
        }
        
        response = self._post(url, data=batch_update_data)
        response.raise_for_status()
        return response.json()

    def list_tools(self):
        return [self.create_document, self.get_document, self.add_content, self.style_text]
