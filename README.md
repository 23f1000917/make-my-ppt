# MakeMyPPT - AI Presentation Generator

MakeMyPPT is a web application that transforms bulk text, markdown, or prose into a styled PowerPoint presentation based on a user-provided template. It uses LLMs (Large Language Models) to structure content and applies the visual style of the uploaded template.

## Features

- **Text Input**: Paste any large text block or markdown content.
- **Style Guidance**: Optional one-line guidance for tone or structure (e.g., "professional deck").
- **Template Support**: Upload a PowerPoint template (.pptx or .potx) to define the visual style.
- **LLM Integration**: Support for multiple LLM providers (OpenAI, Anthropic, Gemini, AI Pipe) using your API keys.
- **Image Reuse**: Reuses images from the template where appropriate (no AI-generated images).
- **Download**: Generates and downloads a new .pptx file.

## How It Works

### Text Parsing and Slide Mapping
1. **Input Processing**: The user's text is sent to an LLM with a prompt that instructs it to break down the content into slides with titles and bullet points.
2. **JSON Structure**: The LLM returns a JSON object containing slide titles and bullet points. Example:
   ```json
   {"slides": [{"title": "Slide Title", "bullets": ["Point 1", "Point 2"]}]}
   ```
3. **Slide Adjustment**: The code enforces minimum and maximum slide counts (default: 10-40 slides) by splitting or merging slides based on bullet point density.

### Template Style Application
1. **Template Analysis**: The uploaded PowerPoint template is parsed using the `python-pptx` library. The app identifies layouts, placeholders, and images.
2. **Style Inheritance**: New slides are created using the template's layout. Titles and body text are placed in appropriate placeholders.
3. **Image Handling**: If enabled, images from the template are reused and placed on slides without overlapping text areas. The app calculates safe zones for image placement.
4. **Visual Consistency**: Fonts, colors, and styles are inherited from the template's theme and layouts. The app attempts to match the template's look and feel as closely as possible.

## Setup and Usage

### Prerequisites
- Python 3.8+
- pip (Python package manager)

### Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/23f1000917/make-my-ppt.git
   cd make-my-ppt
   ```
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the application:
   ```bash
   uvicorn app:app --reload
   ```
4. Open `http://localhost:8000` in your browser.

### Usage
1. **Enter Content**: Paste your text into the content field.
2. **Provide Guidance** (optional): Add style or tone guidance (e.g., "investor pitch").
3. **Upload Template** (optional): Upload a .pptx or .potx file to define the style.
4. **Configure AI Settings**: Select an AI provider, enter your API key, and choose a model.
5. **Generate**: Click "Create Presentation" to generate and download the .pptx file.

## API Key Security
- Your API keys are sent directly to the LLM provider and are not stored or logged by the app.
- All communication with LLM providers is done via secure APIs.

## Code Structure
- `app.py`: FastAPI backend with endpoints for home and presentation creation.
- `index.html`: Frontend interface with HTML, CSS, and JavaScript.

## License
This project is licensed under the MIT License. See the LICENSE file for details.

## Demo
A hosted demo is available at [https://make-my-ppt.vercel.app/].

## Limitations
- Template style inference may not be perfect for complex layouts.
- Image reuse is limited to images from the template; no new images are generated.
- Slide counts are constrained between 10 and 40 by default (adjustable in code).

## Contributing
Contributions are welcome! Please fork the repository and submit pull requests for any improvements.

## Support
For issues or questions, please open an issue on GitHub.
