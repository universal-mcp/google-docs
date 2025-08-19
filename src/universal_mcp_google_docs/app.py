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
        bold: bool = False,
        italic: bool = False,
        underline: bool = False,
        font_size: float | None = None,
        link_url: str | None = None,
        foreground_color: dict[str, float] | None = None,
        background_color: dict[str, float] | None = None,
    ) -> dict[str, Any]:
        """
        Simplified text styling for Google Document - handles most common cases.

        Args:
            document_id: The unique identifier of the Google Document to be updated
            start_index: The zero-based start index of the text range to style
            end_index: The zero-based end index of the text range to style (exclusive)
            bold: Whether the text should be bold
            italic: Whether the text should be italicized
            underline: Whether the text should be underlined
            font_size: The font size in points (e.g., 12.0 for 12pt)
            link_url: URL to make the text a hyperlink
            foreground_color: RGB color dict with 'red', 'green', 'blue' values (0.0 to 1.0)
            background_color: RGB color dict with 'red', 'green', 'blue' values (0.0 to 1.0)

        Returns:
            A dictionary containing the Google Docs API response

        Raises:
            HTTPError: When the API request fails
            RequestException: When there are network connectivity issues

        Tags:
            style, format, text, document, api, google-docs, simple
        """
        url = f"{self.base_api_url}/{document_id}:batchUpdate"
        
        # Build the text style object with only common properties
        text_style = {}
        fields_to_update = []
        
        if bold:
            text_style["bold"] = True
            fields_to_update.append("bold")
        
        if italic:
            text_style["italic"] = True
            fields_to_update.append("italic")
            
        if underline:
            text_style["underline"] = True
            fields_to_update.append("underline")
            
        if font_size is not None:
            text_style["fontSize"] = {"magnitude": font_size, "unit": "PT"}
            fields_to_update.append("fontSize")
            
        if link_url is not None:
            text_style["link"] = {"url": link_url}
            fields_to_update.append("link")
            
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
            fields_to_update.append("foregroundColor")
            
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
            fields_to_update.append("backgroundColor")
        
        # If no styling requested, return early
        if not text_style:
            return {"message": "No styling applied"}
        
        batch_update_data = {
            "requests": [
                {
                    "updateTextStyle": {
                        "range": {
                            "startIndex": start_index,
                            "endIndex": end_index
                        },
                        "textStyle": text_style,
                        "fields": ",".join(fields_to_update)
                    }
                }
            ]
        }
        
        response = self._post(url, data=batch_update_data)
        return self._handle_response(response)

    def delete_content(
        self,
        document_id: str,
        start_index: int,
        end_index: int,
        segment_id: str | None = None,
        tab_id: str | None = None,
    ) -> dict[str, Any]:
        """
        Deletes content from a specified range in a Google Document.

        Args:
            document_id: The unique identifier of the Google Document to be updated
            start_index: The zero-based start index of the content range to delete
            end_index: The zero-based end index of the content range to delete (exclusive)
            segment_id: The ID of the header, footer, or footnote segment (optional)
            tab_id: The ID of the tab containing the content to delete (optional)

        Returns:
            A dictionary containing the Google Docs API response after performing the delete operation

        Raises:
            HTTPError: When the API request fails, such as invalid document_id or insufficient permissions
            RequestException: When there are network connectivity issues or API endpoint problems

        Tags:
            delete, remove, content, document, api, google-docs, batch, content-management, important
        """
        url = f"{self.base_api_url}/{document_id}:batchUpdate"
        
        # Build the delete content range request
        delete_request: dict[str, Any] = {
            "range": {
                "startIndex": start_index,
                "endIndex": end_index
            }
        }
        
        # Add optional parameters if provided
        if segment_id is not None:
            delete_request["range"]["segmentId"] = segment_id
        if tab_id is not None:
            delete_request["tabId"] = tab_id
        
        batch_update_data = {
            "requests": [
                {"deleteContentRange": delete_request}
            ]
        }
        
        response = self._post(url, data=batch_update_data)
        return self._handle_response(response)

    def insert_table(
        self,
        document_id: str,
        location_index: int,
        rows: int,
        columns: int,
        segment_id: str = None,
        tab_id: str = None,
    ) -> dict[str, Any]:
        """
        Inserts a table at the specified location in a Google Document.

        Args:
            document_id: The unique identifier of the Google Document to be updated
            location_index: The zero-based index where the table should be inserted
            rows: The number of rows in the table
            columns: The number of columns in the table
            segment_id: The ID of the header, footer or footnote segment (optional)
            tab_id: The ID of the tab containing the location (optional)

        Returns:
            A dictionary containing the Google Docs API response after performing the insert table operation

        Raises:
            HTTPError: When the API request fails, such as invalid document_id or insufficient permissions
            RequestException: When there are network connectivity issues or API endpoint problems

        Tags:
            table, insert, document, api, google-docs, batch, content-management
        """
        url = f"{self.base_api_url}/{document_id}:batchUpdate"
        
        # Build the location object according to Google Docs API specification
        location = {
            "index": location_index
        }
        
        # Add segment_id if provided (empty string for document body, specific ID for header/footer/footnote)
        if segment_id is not None:
            location["segmentId"] = segment_id
            
        # Add tab_id if provided
        if tab_id is not None:
            location["tabId"] = tab_id
        
        batch_update_data = {
            "requests": [
                {
                    "insertTable": {
                        "location": location,
                        "rows": rows,
                        "columns": columns
                    }
                }
            ]
        }
        
        response = self._post(url, data=batch_update_data)
        return self._handle_response(response)

    def create_footer(
        self,
        document_id: str,
        footer_type: str = "DEFAULT",
        section_break_location_index: int = None,
        section_break_segment_id: str = None,
        section_break_tab_id: str = None,
    ) -> dict[str, Any]:
        """
        Creates a Footer in a Google Document.

        Args:
            document_id: The unique identifier of the Google Document to be updated
            footer_type: The type of footer to create (DEFAULT, HEADER_FOOTER_TYPE_UNSPECIFIED)
            section_break_location_index: The index of the SectionBreak location (optional)
            section_break_segment_id: The segment ID of the SectionBreak location (optional)
            section_break_tab_id: The tab ID of the SectionBreak location (optional)

        Returns:
            A dictionary containing the Google Docs API response after performing the create footer operation

        Raises:
            HTTPError: When the API request fails, such as invalid document_id or insufficient permissions
            RequestException: When there are network connectivity issues or API endpoint problems

        Tags:
            footer, create, document, api, google-docs, batch, content-management
        """
        url = f"{self.base_api_url}/{document_id}:batchUpdate"
        
        # Build the create footer request
        create_footer_request = {
            "type": footer_type
        }
        
        # Add section break location if provided
        if section_break_location_index is not None:
            section_break_location = {
                "index": section_break_location_index
            }
            
            if section_break_segment_id is not None:
                section_break_location["segmentId"] = section_break_segment_id
                
            if section_break_tab_id is not None:
                section_break_location["tabId"] = section_break_tab_id
                
            create_footer_request["sectionBreakLocation"] = section_break_location
        
        batch_update_data = {
            "requests": [
                {"createFooter": create_footer_request}
            ]
        }
        
        response = self._post(url, data=batch_update_data)
        return self._handle_response(response)

    def create_footnote(
        self,
        document_id: str,
        location_index: int = None,
        location_segment_id: str = None,
        location_tab_id: str = None,
        end_of_segment_location: bool = False,
        end_of_segment_segment_id: str = None,
        end_of_segment_tab_id: str = None,
    ) -> dict[str, Any]:
        """
        Creates a Footnote segment and inserts a new FootnoteReference at the given location.

        Args:
            document_id: The unique identifier of the Google Document to be updated
            location_index: The index where to insert the footnote reference (optional)
            location_segment_id: The segment ID for the location (optional, must be empty for body)
            location_tab_id: The tab ID for the location (optional)
            end_of_segment_location: Whether to insert at end of segment (optional)
            end_of_segment_segment_id: The segment ID for end of segment location (optional)
            end_of_segment_tab_id: The tab ID for end of segment location (optional)

        Returns:
            A dictionary containing the Google Docs API response after performing the create footnote operation

        Raises:
            HTTPError: When the API request fails, such as invalid document_id or insufficient permissions
            RequestException: When there are network connectivity issues or API endpoint problems

        Tags:
            footnote, create, document, api, google-docs, batch, content-management
        """
        url = f"{self.base_api_url}/{document_id}:batchUpdate"
        
        # Build the create footnote request
        create_footnote_request = {}
        
        if end_of_segment_location:
            # Use endOfSegmentLocation
            end_of_segment_location_obj = {}
            
            if end_of_segment_segment_id is not None:
                end_of_segment_location_obj["segmentId"] = end_of_segment_segment_id
                
            if end_of_segment_tab_id is not None:
                end_of_segment_location_obj["tabId"] = end_of_segment_tab_id
                
            create_footnote_request["endOfSegmentLocation"] = end_of_segment_location_obj
        else:
            # Use specific location
            location = {
                "index": location_index
            }
            
            if location_segment_id is not None:
                location["segmentId"] = location_segment_id
                
            if location_tab_id is not None:
                location["tabId"] = location_tab_id
                
            create_footnote_request["location"] = location
        
        batch_update_data = {
            "requests": [
                {"createFootnote": create_footnote_request}
            ]
        }
        
        response = self._post(url, data=batch_update_data)
        return self._handle_response(response)

    def delete_footer(
        self,
        document_id: str,
        footer_id: str,
        tab_id: str = None,
    ) -> dict[str, Any]:
        """
        Deletes a Footer from the document.

        Args:
            document_id: The unique identifier of the Google Document to be updated
            footer_id: The ID of the footer to delete
            tab_id: The tab that contains the footer to delete (optional)

        Returns:
            A dictionary containing the Google Docs API response after performing the delete footer operation

        Raises:
            HTTPError: When the API request fails, such as invalid document_id or insufficient permissions
            RequestException: When there are network connectivity issues or API endpoint problems

        Tags:
            footer, delete, remove, document, api, google-docs, batch, content-management
        """
        url = f"{self.base_api_url}/{document_id}:batchUpdate"
        
        # Build the delete footer request
        delete_footer_request = {
            "footerId": footer_id
        }
        
        # Add tab_id if provided
        if tab_id is not None:
            delete_footer_request["tabId"] = tab_id
        
        batch_update_data = {
            "requests": [
                {"deleteFooter": delete_footer_request}
            ]
        }
        
        response = self._post(url, data=batch_update_data)
        return self._handle_response(response)

    def create_header(
        self,
        document_id: str,
        header_type: str = "DEFAULT",
        section_break_location_index: int = None,
        section_break_segment_id: str = None,
        section_break_tab_id: str = None,
    ) -> dict[str, Any]:
        """
        Creates a Header in a Google Document.

        Args:
            document_id: The unique identifier of the Google Document to be updated
            header_type: The type of header to create (DEFAULT, HEADER_FOOTER_TYPE_UNSPECIFIED)
            section_break_location_index: The index of the SectionBreak location (optional)
            section_break_segment_id: The segment ID of the SectionBreak location (optional)
            section_break_tab_id: The tab ID of the SectionBreak location (optional)

        Returns:
            A dictionary containing the Google Docs API response after performing the create header operation

        Raises:
            HTTPError: When the API request fails, such as invalid document_id or insufficient permissions
            RequestException: When there are network connectivity issues or API endpoint problems

        Tags:
            header, create, document, api, google-docs, batch, content-management, important
        """
        url = f"{self.base_api_url}/{document_id}:batchUpdate"
        
        # Build the create header request
        create_header_request = {
            "type": header_type
        }
        
        # Add section break location if provided
        if section_break_location_index is not None:
            section_break_location = {
                "index": section_break_location_index
            }
            
            if section_break_segment_id is not None:
                section_break_location["segmentId"] = section_break_segment_id
                
            if section_break_tab_id is not None:
                section_break_location["tabId"] = section_break_tab_id
                
            create_header_request["sectionBreakLocation"] = section_break_location
        
        batch_update_data = {
            "requests": [
                {"createHeader": create_header_request}
            ]
        }
        
        response = self._post(url, data=batch_update_data)
        return self._handle_response(response)

    def delete_header(
        self,
        document_id: str,
        header_id: str,
        tab_id: str = None,
    ) -> dict[str, Any]:
        """
        Deletes a Header from the document.

        Args:
            document_id: The unique identifier of the Google Document to be updated
            header_id: The ID of the header to delete
            tab_id: The tab containing the header to delete (optional)

        Returns:
            A dictionary containing the Google Docs API response after performing the delete header operation

        Raises:
            HTTPError: When the API request fails, such as invalid document_id or insufficient permissions
            RequestException: When there are network connectivity issues or API endpoint problems

        Tags:
            header, delete, remove, document, api, google-docs, batch, content-management
        """
        url = f"{self.base_api_url}/{document_id}:batchUpdate"
        
        # Build the delete header request
        delete_header_request = {
            "headerId": header_id
        }
        
        # Add tab_id if provided
        if tab_id is not None:
            delete_header_request["tabId"] = tab_id
        
        batch_update_data = {
            "requests": [
                {"deleteHeader": delete_header_request}
            ]
        }
        
        response = self._post(url, data=batch_update_data)
        return self._handle_response(response)

    def create_paragraph_bullets(
        self,
        document_id: str,
        start_index: int,
        end_index: int,
        bullet_preset: str,
        segment_id: str = None,
        tab_id: str = None,
    ) -> dict[str, Any]:
        """
        Creates bullets for all of the paragraphs that overlap with the given range.

        Args:
            document_id: The unique identifier of the Google Document to be updated
            start_index: The zero-based start index of the range to apply bullets to
            end_index: The zero-based end index of the range to apply bullets to (exclusive)
            bullet_preset: The kind of bullet glyphs to use. Available options:
                - BULLET_GLYPH_PRESET_UNSPECIFIED: The bullet glyph preset is unspecified
                - BULLET_DISC_CIRCLE_SQUARE: DISC, CIRCLE and SQUARE for first 3 nesting levels
                - BULLET_DIAMONDX_ARROW3D_SQUARE: DIAMONDX, ARROW3D and SQUARE for first 3 nesting levels
                - BULLET_CHECKBOX: CHECKBOX bullet glyphs for all nesting levels
                - BULLET_ARROW_DIAMOND_DISC: ARROW, DIAMOND and DISC for first 3 nesting levels
                - BULLET_STAR_CIRCLE_SQUARE: STAR, CIRCLE and SQUARE for first 3 nesting levels
                - BULLET_ARROW3D_CIRCLE_SQUARE: ARROW3D, CIRCLE and SQUARE for first 3 nesting levels
                - BULLET_LEFTTRIANGLE_DIAMOND_DISC: LEFTTRIANGLE, DIAMOND and DISC for first 3 nesting levels
                - BULLET_DIAMONDX_HOLLOWDIAMOND_SQUARE: DIAMONDX, HOLLOWDIAMOND and SQUARE for first 3 nesting levels
                - BULLET_DIAMOND_CIRCLE_SQUARE: DIAMOND, CIRCLE and SQUARE for first 3 nesting levels
                - NUMBERED_DECIMAL_ALPHA_ROMAN: DECIMAL, ALPHA and ROMAN with periods
                - NUMBERED_DECIMAL_ALPHA_ROMAN_PARENS: DECIMAL, ALPHA and ROMAN with parenthesis
                - NUMBERED_DECIMAL_NESTED: DECIMAL with nested numbering (1., 1.1., 2., 2.2.)
                - NUMBERED_UPPERALPHA_ALPHA_ROMAN: UPPERALPHA, ALPHA and ROMAN with periods
                - NUMBERED_UPPERROMAN_UPPERALPHA_DECIMAL: UPPERROMAN, UPPERALPHA and DECIMAL with periods
                - NUMBERED_ZERODECIMAL_ALPHA_ROMAN: ZERODECIMAL, ALPHA and ROMAN with periods
            segment_id: The segment ID for the range (optional)
            tab_id: The tab ID for the range (optional)

        Returns:
            A dictionary containing the Google Docs API response after performing the create bullets operation

        Raises:
            HTTPError: When the API request fails, such as invalid document_id or insufficient permissions
            RequestException: When there are network connectivity issues or API endpoint problems

        Tags:
            bullets, list, paragraph, document, api, google-docs, batch, content-management
        """
        url = f"{self.base_api_url}/{document_id}:batchUpdate"
        
        # Build the range object
        range_obj = {
            "startIndex": start_index,
            "endIndex": end_index
        }
        
        # Add optional parameters if provided
        if segment_id is not None:
            range_obj["segmentId"] = segment_id
        if tab_id is not None:
            range_obj["tabId"] = tab_id
        
        batch_update_data = {
            "requests": [
                {
                    "createParagraphBullets": {
                        "range": range_obj,
                        "bulletPreset": bullet_preset
                    }
                }
            ]
        }
        
        response = self._post(url, data=batch_update_data)
        return self._handle_response(response)

    def delete_paragraph_bullets(
        self,
        document_id: str,
        start_index: int,
        end_index: int,
        segment_id: str = None,
        tab_id: str = None,
    ) -> dict[str, Any]:
        """
        Deletes bullets from all of the paragraphs that overlap with the given range.

        Args:
            document_id: The unique identifier of the Google Document to be updated
            start_index: The zero-based start index of the range to remove bullets from
            end_index: The zero-based end index of the range to remove bullets from (exclusive)
            segment_id: The segment ID for the range (optional)
            tab_id: The tab ID for the range (optional)

        Returns:
            A dictionary containing the Google Docs API response after performing the delete bullets operation

        Raises:
            HTTPError: When the API request fails, such as invalid document_id or insufficient permissions
            RequestException: When there are network connectivity issues or API endpoint problems

        Tags:
            bullets, delete, remove, list, paragraph, document, api, google-docs, batch, content-management
        """
        url = f"{self.base_api_url}/{document_id}:batchUpdate"
        
        # Build the range object
        range_obj = {
            "startIndex": start_index,
            "endIndex": end_index
        }
        
        # Add optional parameters if provided
        if segment_id is not None:
            range_obj["segmentId"] = segment_id
        if tab_id is not None:
            range_obj["tabId"] = tab_id
        
        batch_update_data = {
            "requests": [
                {
                    "deleteParagraphBullets": {
                        "range": range_obj
                    }
                }
            ]
        }
        
        response = self._post(url, data=batch_update_data)
        return self._handle_response(response)

    def list_tools(self):
        return [self.create_document, self.get_document, self.add_content, self.style_text, self.delete_content, self.insert_table, self.create_footer, self.create_footnote, self.delete_footer, self.create_header, self.delete_header, self.create_paragraph_bullets, self.delete_paragraph_bullets]
