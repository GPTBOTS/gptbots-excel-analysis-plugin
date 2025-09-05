# Excel AI Plugin - Simple Guide

## What is it?

An Excel add-in (XLAM) that adds AI analysis capabilities to your spreadsheets. Ask questions about your data and get intelligent insights directly in Excel.

## Key Features

- **AI Analysis**: Ask questions about your spreadsheet data
- **Easy Menu**: Adds "AI工具" button to Excel menu bar
- **Smart Results**: Creates new worksheet with formatted AI responses
- **Markdown Support**: Renders formatted text with headers, bold, code blocks

## How to Install

1. Download the XLAM file
2. Open Excel → File → Options → Add-ins
3. Click "Go" → "Browse" → Select XLAM file
4. Check the box and click "OK"
5. Restart Excel

## How to Use

1. Open Excel with your data
2. Click "AI工具" in the menu bar
3. Type your question (e.g., "What are the trends in this data?")
4. Wait for AI analysis
5. Review results in the new worksheet

## Example Questions

- "Summarize this data"
- "Find patterns in the sales numbers"
- "What are the key insights?"
- "Identify any problems in this data"
- "Create a report of the top performers"

## Technical Details

### What it does:
- Extracts all data from your active worksheet
- Sends data + your question to AI API
- Formats the AI response with proper styling
- Creates a new worksheet with the results

### API Integration:
- Endpoint: `https://api.xxxx.ai/xx/conversation/xxxxxxx`
- Method: POST with JSON data
- Authentication: Bearer token
- Response: JSON with text content

### Data Format:
- Input: Tab-separated worksheet data
- Output: Formatted text with markdown rendering
- Styling: Headers, bold text, code blocks, colors

## Error Handling

- **No Internet**: Shows "无法创建HTTP对象" error
- **Empty Response**: Shows "API无返回结果" message
- **Empty Question**: Prompts to enter a question
- **API Errors**: Displays error messages

## Requirements

- Excel 2016 or later
- Windows OS
- Internet connection
- Valid API access token

## Security

- Data sent to AI API via HTTPS
- No local data storage
- User controls when data is sent
- Bearer token authentication

## Troubleshooting

**Menu not showing?**
- Check if add-in is enabled in Excel options
- Restart Excel after installation

**API not working?**
- Check internet connection
- Verify API endpoint and token
- Check firewall settings

**No results?**
- Look for new worksheet with timestamp name
- Check for error messages
- Verify API response format

## Customization

You can modify:
- API endpoint URL
- Authentication token
- Menu text
- Response formatting
- Error messages

## Support

For issues or questions:
- Check this documentation
- Review VBA code comments
- Contact technical support

## Version Info

- Current: v1.0.0
- Features: Basic AI integration, markdown rendering, menu interface
- Compatibility: Excel 2016+

---

**Note**: This plugin requires an active AI API service and valid authentication credentials to function properly.
