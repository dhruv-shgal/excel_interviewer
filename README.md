# Excel Mock Interviewer 📊

An AI-powered Streamlit application that simulates Excel interviews with progressive difficulty and real-time feedback.

## LINK - https://excel-interviewer.streamlit.app/

## Features

- **Random Question Generation**: Dynamically creates Excel interview questions with progressive difficulty
- **Real-time Evaluation**: Provides instant feedback on your answers
- **Session Management**: Tracks your progress through a complete interview
- **Professional Reports**: Generates detailed feedback and recommendations
- **Fallback System**: Automatically uses predefined questions when API limits are reached
- **Clean UI**: Simple, markdown-based interface for easy use

## Topics Covered

- VLOOKUP and HLOOKUP functions
- INDEX and MATCH functions
- IF statements and nested logic
- Pivot tables and data analysis
- Data cleaning and validation
- Conditional formatting
- Array formulas
- Financial functions
- Statistical functions
- Text manipulation functions
- Date and time functions
- Data visualization and charts

## Installation & Setup

### Local Development

1. **Clone or download the project files**
   ```bash
   # Ensure you have the following files:
   # - main.py
   # - requirements.txt
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application**
   ```bash
   streamlit run main.py
   ```

4. **Open your browser**
   - The app will automatically open at `http://localhost:8501`
   - If not, manually navigate to the URL shown in the terminal

### Hugging Face Spaces Deployment

1. **Create a new Space**
   - Go to [Hugging Face Spaces](https://huggingface.co/spaces)
   - Click "Create new Space"
   - Choose "Streamlit" as the SDK
   - Set visibility (Public/Private)

2. **Upload files**
   - Upload `main.py` and `requirements.txt` to your Space
   - The app will automatically deploy

3. **Configure secrets (if needed)**
   - The Hugging Face token is already included in the code
   - For production, consider using Space secrets for the token

### Streamlit Cloud Deployment

1. **Push to GitHub**
   - Create a GitHub repository
   - Upload `main.py` and `requirements.txt`

2. **Deploy on Streamlit Cloud**
   - Go to [share.streamlit.io](https://share.streamlit.io)
   - Connect your GitHub repository
   - Deploy the app

## Usage

1. **Start the Interview**
   - Click "Start Interview" on the welcome screen
   - The AI will generate your first question

2. **Answer Questions**
   - Read each question carefully
   - Provide detailed, specific answers
   - Click "Submit & Evaluate" when ready

3. **Review Feedback**
   - Get immediate scoring (0-2) for each answer
   - Read the AI-generated feedback
   - Progress through all 5 questions

4. **Final Report**
   - Receive a comprehensive performance summary
   - View strengths and areas for improvement
   - Get your final score out of 10

## API Configuration

The app uses an external AI API for question generation and answer evaluation. The API key is included in the code for easy setup.

**For production deployment, consider:**
- Using environment variables for the API key
- Implementing rate limiting
- Using the built-in fallback mechanism when API limits are reached

## Sample Interview Runs

The code includes commented examples of:
- **Good Candidate**: Scores 10/10 with comprehensive answers
- **Weak Candidate**: Scores 2/10 with basic responses

## Technical Details

- **Framework**: Streamlit
- **AI Integration**: External AI API for question generation and evaluation
- **State Management**: Streamlit session state
- **Styling**: Clean Markdown formatting with minimal styling
- **Dependencies**: See requirements.txt

## Fallback System

The application includes a robust fallback mechanism to ensure uninterrupted functionality:

- **Automatic Detection**: Detects when API limits are reached or API errors occur
- **Predefined Questions**: Uses a curated set of Excel interview questions organized by difficulty level
- **Seamless Transition**: Switches to fallback questions without disrupting the user experience
- **Progressive Difficulty**: Maintains the same difficulty progression as the AI-generated questions

## Troubleshooting

**Common Issues:**

1. **API Errors**: Check your internet connection and API service status
2. **API Limits**: The application will automatically fall back to predefined questions when API limits are reached
3. **Empty Answers**: Make sure to provide detailed answers before submitting

**Performance Tips:**
- Use a stable internet connection
- Provide detailed answers for better evaluation
- If you encounter API limits, the fallback system will ensure the application continues to function

## Contributing

Feel free to submit issues, feature requests, or pull requests to improve the application.

## License

This project is open source and available under the MIT License.
