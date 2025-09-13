import streamlit as st
import requests
import json
import random
from typing import Dict, List, Tuple
import google.generativeai as genai
import json


st.set_page_config(
    page_title="Excel Mock Interviewer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)



GEMINI_API_KEY = "YOUR_API_KEY"
genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel('gemini-1.5-flash')




FALLBACK_QUESTIONS = [
    "How would you use VLOOKUP to find the price of a product based on its ID, and what are the key parameters you need to consider?",
    "Explain the difference between VLOOKUP and INDEX/MATCH functions. When would you use each one?",
    "How would you create a pivot table to analyze sales data by region and product category?",
    "Describe the process of cleaning data with duplicate entries and inconsistent formatting in Excel.",
    "How would you use conditional formatting to highlight cells that meet multiple criteria simultaneously?",
    "Explain how to create an array formula that calculates the sum of values based on multiple conditions.",
    "How would you use the IF function with nested logic to categorize data into different performance levels?",
    "Describe how to create a dynamic dashboard using pivot tables, slicers, and charts.",
    "How would you use the INDEX and MATCH functions together to perform a two-way lookup?",
    "Explain the process of creating data validation rules to ensure data integrity in Excel.",
    "How would you use the PMT function to calculate monthly loan payments and what parameters are required?",
    "Describe how to use the CONCATENATE function and text functions to clean and format data.",
    "How would you create a chart that automatically updates when new data is added to your worksheet?",
    "Explain how to use the SUMIFS function to sum values based on multiple criteria.",
    "How would you create a macro to automate repetitive tasks in Excel?"
]


EXCEL_TOPICS = [
    "VLOOKUP and HLOOKUP functions",
    "INDEX and MATCH functions",
    "IF statements and nested IFs",
    "Pivot tables and data analysis",
    "Data cleaning and validation",
    "Conditional formatting",
    "Array formulas",
    "Data visualization with charts",
    "Advanced filtering and sorting",
    "Macros and automation",
    "Financial functions (PMT, NPV, IRR)",
    "Statistical functions (AVERAGEIF, COUNTIF, SUMIF)",
    "Text functions (CONCATENATE, LEFT, RIGHT, MID)",
    "Date and time functions",
    "Lookup and reference functions"
]

def call_gemini_api(prompt: str, max_length: int = 500) -> str:
    """
    Call Google Gemini API for question generation and evaluation.
    """
    try:
        response = model.generate_content(
            prompt,
            generation_config=genai.types.GenerationConfig(
                max_output_tokens=max_length,
                temperature=0.7,
            )
        )
        return response.text.strip()
    
    except Exception as e:
        
        return None

def generate_excel_question(question_number: int) -> str:
    """
    Generate an Excel interview question with progressive difficulty using Google Gemini.
    """
 
    difficulty_levels = {
        1: {
            "level": "Basic",
            "topics": ["Basic formulas", "Simple functions", "Data entry", "Basic formatting"],
            "description": "Beginner Excel skills - basic formulas, simple functions, data entry"
        },
        2: {
            "level": "Intermediate-Basic", 
            "topics": ["VLOOKUP", "IF statements", "Basic pivot tables", "Data sorting"],
            "description": "Intermediate Excel skills - lookup functions, conditional logic, basic analysis"
        },
        3: {
            "level": "Intermediate",
            "topics": ["INDEX/MATCH", "Pivot tables", "Data cleaning", "Charts"],
            "description": "Intermediate Excel skills - advanced lookups, data analysis, visualization"
        },
        4: {
            "level": "Advanced-Intermediate",
            "topics": ["Array formulas", "Conditional formatting", "Data validation", "Advanced pivot tables"],
            "description": "Advanced Excel skills - complex formulas, data validation, advanced analysis"
        },
        5: {
            "level": "Advanced",
            "topics": ["Macros", "Financial functions", "Dashboard creation", "Advanced automation"],
            "description": "Expert Excel skills - automation, complex analysis, professional dashboards"
        }
    }
    
    current_level = difficulty_levels[question_number]
    topic = random.choice(current_level["topics"])
    
    prompt = f"""Generate a specific, practical Excel interview question for a {current_level['level']} level candidate.
    
    Difficulty Level: {current_level['level']}
    Focus Topic: {topic}
    Description: {current_level['description']}
    
    The question should:
    - Match the {current_level['level']} difficulty level
    - Be realistic and job-relevant
    - Test appropriate Excel skills for this level
    - Require a detailed answer
    - Be clear and specific
    - Progress naturally from basic to advanced concepts
    
    Format: Just provide the question, no additional text.
    
    Examples by level:
    - Basic: "How would you calculate the sum of values in column A using a formula?"
    - Intermediate: "How would you use VLOOKUP to find a product price based on its ID?"
    - Advanced: "How would you create a dynamic dashboard with pivot tables and slicers?"
    
    Generate a {current_level['level']} level question about {topic}:"""
    
    
    ai_question = call_gemini_api(prompt, max_length=200)
    
  
    if ai_question is None or "Error" in ai_question:
    
        return get_fallback_question_by_difficulty(question_number)
    
    return ai_question

def get_fallback_question_by_difficulty(question_number: int) -> str:
    """
    Get fallback questions organized by difficulty level.
    """
    fallback_questions_by_level = {
        1: [
            "How would you calculate the total sales for a month using a SUM formula?",
            "Explain how to use the AVERAGE function to find the average of a range of numbers.",
            "How would you format cells to display currency values with dollar signs?",
            "What is the difference between relative and absolute cell references in Excel?"
        ],
        2: [
            "How would you use VLOOKUP to find the price of a product based on its ID?",
            "Explain how to use the IF function to categorize data into different performance levels.",
            "How would you create a basic pivot table to summarize sales data by region?",
            "How would you sort data by multiple columns in Excel?"
        ],
        3: [
            "How would you use INDEX and MATCH functions together to perform a two-way lookup?",
            "Explain how to clean data with duplicate entries and inconsistent formatting.",
            "How would you create a chart that automatically updates when new data is added?",
            "How would you use conditional formatting to highlight cells based on specific criteria?"
        ],
        4: [
            "How would you create an array formula to calculate the sum of values based on multiple conditions?",
            "Explain how to set up data validation rules to ensure data integrity in Excel.",
            "How would you create a dynamic dashboard using pivot tables, slicers, and charts?",
            "How would you use advanced conditional formatting with custom formulas?"
        ],
        5: [
            "How would you create a macro to automate repetitive data entry tasks?",
            "Explain how to use financial functions like PMT, FV, and NPV for loan calculations.",
            "How would you build a comprehensive dashboard with multiple data sources and interactive elements?",
            "How would you implement advanced data analysis using Power Query and Power Pivot?"
        ]
    }
    
    return random.choice(fallback_questions_by_level[question_number])

def evaluate_answer(question: str, answer: str) -> Tuple[int, str]:
    """
    Evaluate the candidate's answer and return a score (0-2) and explanation.
    """
    prompt = f"""Evaluate this Excel interview answer and provide a score and brief explanation.

    Question: {question}
    Answer: {answer}
    
    Scoring criteria:
    - 0: Incorrect or completely wrong approach
    - 1: Partially correct but missing key elements or has errors
    - 2: Correct and comprehensive answer
    
    Respond in this exact format:
    Score: [0, 1, or 2]
    Explanation: [One-line explanation of the score]
    
    Be fair but thorough in your evaluation. Focus on technical accuracy and completeness."""
    
    response = call_gemini_api(prompt, max_length=150)
    
 
    if response is None or "Error" in response:
  
        answer_lower = answer.lower()
        excel_keywords = [
            'vlookup', 'hlookup', 'index', 'match', 'pivot', 'pivot table', 
            'formula', 'function', 'conditional formatting', 'data validation',
            'array formula', 'sumif', 'countif', 'averageif', 'sumifs', 'countifs',
            'if statement', 'nested if', 'concatenate', 'text functions',
            'pmt', 'fv', 'pv', 'financial functions', 'date functions',
            'chart', 'dashboard', 'slicer', 'filter', 'sort', 'macro'
        ]
        
        keyword_count = sum(1 for keyword in excel_keywords if keyword in answer_lower)
        
        if len(answer) < 20:
            return 0, "Answer too brief - please provide more detail"
        elif keyword_count >= 3 and len(answer) > 100:
            return 2, "Excellent technical knowledge demonstrated"
        elif keyword_count >= 2 and len(answer) > 50:
            return 2, "Good technical knowledge with relevant Excel concepts"
        elif keyword_count >= 1 and len(answer) > 30:
            return 1, "Shows some Excel knowledge but could be more detailed"
        elif any(word in answer_lower for word in ['excel', 'spreadsheet', 'data', 'table']):
            return 1, "Basic understanding shown but needs more technical detail"
        else:
            return 0, "Limited Excel knowledge demonstrated"
    

    try:
        lines = response.split('\n')
        score = 0
        explanation = "No explanation provided"
        
        for line in lines:
            if line.startswith("Score:"):
                score_text = line.replace("Score:", "").strip()
                if "0" in score_text:
                    score = 0
                elif "1" in score_text:
                    score = 1
                elif "2" in score_text:
                    score = 2
            elif line.startswith("Explanation:"):
                explanation = line.replace("Explanation:", "").strip()
        
        return score, explanation
    
    except Exception:
        return 1, "Evaluation error - partial credit given"

def evaluate_all_answers(questions_answers: List[Dict]) -> List[Dict]:
    """
    Evaluate all answers at once when the user submits all answers.
    """
    with st.spinner("Evaluating all your answers... This may take a moment."):
        for i, qa in enumerate(questions_answers):
      
            if i in st.session_state.skipped_questions:
                qa['score'] = 0
                qa['explanation'] = 'Question was skipped'
            elif qa['score'] == 0 and qa['explanation'] == '':
      
                score, explanation = evaluate_answer(qa['question'], qa['answer'])
                qa['score'] = score
                qa['explanation'] = explanation
    
    return questions_answers

def generate_final_report(questions_answers: List[Dict]) -> str:
    """
    Generate a professional feedback report using Google Gemini.
    """

    qa_summary = ""
    total_score = 0
    

    correct_answers = {
        0: "To calculate the total cost, you would use the SUMPRODUCT function. In cell C7, the formula would be =SUMPRODUCT(A2:A6,B2:B6) which multiplies each item's price by its quantity and then adds all the results together.",
        1: "To create a dynamic chart that updates automatically, you would: 1) Create a named range for your data (Ctrl+T or Insert > Table), 2) Insert a chart based on this table (Insert > Charts > desired chart type), 3) The chart will automatically update when data in the table changes. You can also use OFFSET or INDEX functions with COUNTA to create dynamic ranges.",
        2: "To find the last value in column A, you can use: =LOOKUP(2,1/(A:A<>""),A:A) or =INDEX(A:A,MATCH(9.99999999999999E+307,A:A)) or =INDEX(A:A,COUNTA(A:A)). These formulas work even when the data has blank cells or is unsorted.",
        3: "To create a conditional formatting rule that highlights cells with values above average: 1) Select the range, 2) Go to Home > Conditional Formatting > New Rule, 3) Choose 'Use a formula', 4) Enter =A1>AVERAGE($A$1:$A$100) (adjust range as needed), 5) Click Format and choose highlighting style, 6) Click OK. This will highlight all cells with values above the average of the selected range.",
        4: "To create a pivot table summarizing sales by region and product: 1) Select your data range, 2) Go to Insert > PivotTable, 3) In the PivotTable Fields pane, drag 'Region' to Rows area, 'Product' to Columns area, and 'Sales' to Values area, 4) The pivot table will automatically calculate the sum of sales for each region-product combination. You can then add filters, change calculation type (e.g., to average), or add additional fields as needed."
    }
    
    for i, qa in enumerate(questions_answers, 1):
        qa_summary += f"Q{i}: {qa['question']}\n"
        qa_summary += f"Your Answer: {qa['answer']}\n"
        qa_summary += f"Correct Answer: {correct_answers[i-1]}\n"
        qa_summary += f"Score: {qa['score']}/2\n\n"
        total_score += qa['score']
    

    percentage = (total_score / 10) * 100
    

    if percentage >= 90:
        skill_level = "Expert"
    elif percentage >= 75:
        skill_level = "Advanced"
    elif percentage >= 60:
        skill_level = "Intermediate"
    elif percentage >= 40:
        skill_level = "Basic"
    else:
        skill_level = "Beginner"
    
    prompt = f"""Generate a professional interview feedback report for an Excel mock interview.

    Interview Summary:
    {qa_summary}
    
    Total Score: {total_score}/10 ({percentage:.1f}%), which puts the candidate at a {skill_level} level.
    
    Create a professional feedback report with these sections:
    1. Overall Performance Summary - include specific strengths demonstrated
    2. Strengths (areas where candidate performed well) - be specific about which Excel skills they demonstrated proficiency in
    3. Areas for Improvement (weaknesses) - provide detailed analysis based on questions they struggled with
    4. Specific Recommendations - include actionable steps and suggested resources tailored to their skill gaps
    5. Final Score and Recommendation - with a motivational conclusion that encourages continued learning
    
    Be constructive, specific, and professional. Focus on Excel skills and provide actionable feedback that will help them improve their Excel skills for real-world applications.
    
    Important formatting guidelines:
    - Do not include any placeholder text like [Candidate Name] or [Interviewer Name]
    - Do not use asterisks (*) for formatting - use proper markdown formatting instead
    - Focus only on genuine evaluation terms and feedback
    - Make sure all sections are properly formatted with clear headings"""
    
    response = call_gemini_api(prompt, max_length=500)
    
    if response is None or "Error" in response:

        high_scores = [qa for qa in questions_answers if qa['score'] == 2]
        low_scores = [qa for qa in questions_answers if qa['score'] == 0]
        
        strengths = []
        if high_scores:
            strengths.append(f"Strong performance on {len(high_scores)} questions with detailed technical knowledge")
        if any('vlookup' in qa['answer'].lower() or 'index' in qa['answer'].lower() for qa in high_scores):
            strengths.append("Good understanding of lookup functions and reference techniques")
        if any('pivot' in qa['answer'].lower() for qa in high_scores):
            strengths.append("Solid grasp of pivot tables and data analysis capabilities")
        if any('formula' in qa['answer'].lower() for qa in high_scores):
            strengths.append("Strong formula and function knowledge with practical application")
        if any('chart' in qa['answer'].lower() or 'dashboard' in qa['answer'].lower() for qa in high_scores):
            strengths.append("Effective data visualization and dashboard creation skills")
        
 
        improvements = []
        if low_scores:
            improvements.append(f"Focus on {len(low_scores)} areas where technical knowledge needs strengthening")
        if any('vlookup' in qa['question'].lower() and qa['score'] < 2 for qa in questions_answers):
            improvements.append("Practice with VLOOKUP, INDEX/MATCH and advanced lookup techniques")
        if any('pivot' in qa['question'].lower() and qa['score'] < 2 for qa in questions_answers):
            improvements.append("Review pivot table creation, calculated fields, and advanced data analysis techniques")
        if any('formula' in qa['question'].lower() and qa['score'] < 2 for qa in questions_answers):
            improvements.append("Strengthen formula writing, nested functions, and array formula usage")
        if any('chart' in qa['question'].lower() and qa['score'] < 2 for qa in questions_answers):
            improvements.append("Improve data visualization techniques and interactive dashboard creation")
        
        if total_score >= 8:
            recommendation = f"Excellent performance at {skill_level} level! You demonstrate strong Excel skills suitable for advanced roles."
            resources = "- Microsoft's official Power BI and Power Query documentation\n- Advanced Excel courses on LinkedIn Learning or Coursera\n- Excel MVP blogs and forums for cutting-edge techniques"
        elif total_score >= 5:
            recommendation = f"Good job at {skill_level} level! You have solid Excel knowledge with room for improvement in specific areas."
            resources = "- ExcelJet.net for advanced function tutorials\n- Chandoo.org for practical Excel applications\n- YouTube channels like ExcelIsFun for detailed walkthroughs"
        else:
            recommendation = f"You're currently at {skill_level} level. Consider additional Excel training and practice to strengthen your technical skills."
            resources = "- Microsoft's Excel Essentials training\n- GCF Learn Free Excel tutorials\n- YouTube channels like Leila Gharani for beginner-friendly guides"
        
        return f"""**Professional Feedback Report**

**Overall Performance Summary:**
You completed the Excel mock interview with a total score of {total_score}/10 ({percentage:.1f}%), which places you at a {skill_level} level.

**Strengths:**
{chr(10).join(f"- {strength}" for strength in strengths) if strengths else "- Showed effort in completing all questions"}

**Areas for Improvement:**
{chr(10).join(f"- {improvement}" for improvement in improvements) if improvements else "- Focus on Excel functions and advanced features"}

**Final Score and Recommendation:**
Total Score: {total_score}/10
Recommendation: {recommendation}

**Suggested Resources:**
{resources}

**Next Steps:**
- Practice with Excel functions and formulas regularly with real-world datasets
- Work on pivot tables and data analysis techniques through guided projects
- Review conditional formatting and data validation for data integrity
- Consider taking advanced Excel courses for professional development
- Apply your skills to solve actual business problems to reinforce learning

Remember that Excel proficiency comes with consistent practice. Keep challenging yourself with new problems and techniques!"""
    
    return response

def initialize_session_state():
    """Initialize session state variables if they don't exist."""
    if 'interview_started' not in st.session_state:
        st.session_state.interview_started = False
    if 'current_question' not in st.session_state:
        st.session_state.current_question = 0
    if 'questions_answers' not in st.session_state:
        st.session_state.questions_answers = []
    if 'current_answer' not in st.session_state:
        st.session_state.current_answer = ""
    if 'interview_complete' not in st.session_state:
        st.session_state.interview_complete = False
    if 'show_submit_all' not in st.session_state:
        st.session_state.show_submit_all = False
    if 'answers_evaluated' not in st.session_state:
        st.session_state.answers_evaluated = False
    if 'marked_for_review' not in st.session_state:
        st.session_state.marked_for_review = [False, False, False, False, False]
    if 'skipped_questions' not in st.session_state:
        st.session_state.skipped_questions = []

# API test function removed - using self-contained logic

def display_interview_introduction():
    """Display the interview introduction and agent introduction."""
    st.markdown("""
    <div style="background-color: #1e3a8a; padding: 2rem; border-radius: 10px; margin: 1rem 0;">
        <h2 style="color: #ffffff; text-align: center; margin-bottom: 1rem;">🤖 AI Interview Agent</h2>
        <p style="color: #e2e8f0; text-align: center; font-size: 1.1rem;">
            Hello! I'm your AI Excel Interview Agent. I'll be conducting a structured interview to assess your Excel skills.
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("""
    ## 📋 Interview Process Overview
    
    I'll guide you through a **5-question progressive interview** that will:
    
    - **Start with basic concepts** and gradually increase in complexity
    - **Test practical Excel skills** relevant to real-world scenarios
    - **Provide immediate feedback** after each question
    - **Generate a comprehensive report** with strengths and improvement areas
    
    **Interview Structure:**
    1. 🟢 **Basic Level** - Fundamental Excel operations
    2. 🟡 **Intermediate-Basic** - Lookup functions and basic analysis
    3. 🟠 **Intermediate** - Data analysis and visualization
    4. 🔴 **Advanced-Intermediate** - Complex formulas and validation
    5. 🟣 **Advanced** - Automation and professional dashboards
    
    **Evaluation Criteria:**
    - Technical accuracy and completeness
    - Practical application of Excel features
    - Problem-solving approach
    - Depth of understanding
    
    Are you ready to begin the interview?
    """)

def display_agent_thinking(question_number: int):
    """Display agent thinking process."""
    difficulty_levels = {
        1: {"level": "Basic", "color": "🟢", "description": "Let me start with fundamental Excel concepts..."},
        2: {"level": "Intermediate-Basic", "color": "🟡", "description": "Now I'll test your lookup and analysis skills..."},
        3: {"level": "Intermediate", "color": "🟠", "description": "Time to evaluate your data analysis capabilities..."},
        4: {"level": "Advanced-Intermediate", "color": "🔴", "description": "Let's see how you handle complex formulas..."},
        5: {"level": "Advanced", "color": "🟣", "description": "Final challenge - advanced automation and dashboards..."}
    }
    
    current_level = difficulty_levels[question_number]
    
    st.markdown(f"""
    <div style="background-color: #374151; padding: 1rem; border-radius: 8px; margin: 1rem 0;">
        <p style="color: #e5e7eb; margin: 0;">
            <strong>🤖 Agent Thinking:</strong> {current_level['description']}
        </p>
    </div>
    """, unsafe_allow_html=True)

def display_agent_evaluation():
    """Display agent evaluation process."""
    st.markdown("""
    <div style="background-color: #374151; padding: 1rem; border-radius: 8px; margin: 1rem 0;">
        <p style="color: #e5e7eb; margin: 0;">
            <strong>🤖 Agent Evaluating:</strong> Analyzing your response for technical accuracy, completeness, and practical application...
        </p>
    </div>
    """, unsafe_allow_html=True)

def main():
    """Main Streamlit application with structured interview flow."""
    initialize_session_state()
    

    st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .question-box {
        background-color: #2d3748;
        color: #ffffff;
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #1f77b4;
        margin: 1rem 0;
        border: 1px solid #4a5568;
    }
    .answer-box {
        background-color: #2d3748;
        color: #ffffff;
        padding: 1rem;
        border-radius: 8px;
        margin: 1rem 0;
        border: 1px solid #4a5568;
    }
    .score-box {
        background-color: #2d3748;
        color: #ffffff;
        padding: 1rem;
        border-radius: 8px;
        margin: 1rem 0;
        border: 1px solid #4a5568;
    }
    .report-box {
        background-color: #2d3748;
        color: #ffffff;
        padding: 2rem;
        border-radius: 10px;
        border: 2px solid #1f77b4;
        margin: 2rem 0;
        box-shadow: 0 6px 12px rgba(0, 0, 0, 0.2);
    }
    .agent-message {
        background-color: #1e3a8a;
        color: #ffffff;
        padding: 1rem;
        border-radius: 8px;
        margin: 1rem 0;
        border-left: 4px solid #3b82f6;
    }
    .submit-all-btn {
        background-color: #1f77b4;
        color: white;
        font-weight: bold;
        padding: 0.75rem;
        border-radius: 8px;
        text-align: center;
        margin: 1rem 0;
        cursor: pointer;
    }
    .stButton>button {
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #45a049;
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2);
    }
    .submit-all-button button {
        background-color: #2196F3 !important;
        color: white !important;
        font-weight: bold !important;
        font-size: 1.1rem !important;
        padding: 0.5rem 1rem !important;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1) !important;
    }
    .submit-all-button button:hover {
        background-color: #0b7dda !important;
        box-shadow: 0 6px 10px rgba(0, 0, 0, 0.2) !important;
    }
    .question-item {
        border-bottom: 1px solid #4a5568;
        padding-bottom: 15px;
        margin-bottom: 15px;
    }
    .question-text {
        font-weight: 600;
        color: #e2e8f0;
        margin-bottom: 10px;
    }
    .user-answer {
        background-color: #1a202c;
        padding: 10px;
        border-radius: 5px;
        margin-bottom: 10px;
        border-left: 3px solid #4a5568;
    }
    .correct-answer {
        background-color: #1a3e2c;
        padding: 10px;
        border-radius: 5px;
        border-left: 4px solid #4caf50;
        margin-bottom: 10px;
    }
    .score-indicator {
        font-weight: bold;
        padding: 3px 8px;
        border-radius: 12px;
        display: inline-block;
        margin-left: 10px;
    }
    .score-0 {
        background-color: #5c1e1e;
        color: #ff8a8a;
    }
    .score-1 {
        background-color: #5c4d1e;
        color: #ffd78a;
    }
    .score-2 {
        background-color: #1e5c2f;
        color: #8aff8a;
    }
    .report-section {
        margin-bottom: 20px;
        border-left: 3px solid #3b82f6;
        padding-left: 15px;
    }
    .report-section h3 {
        border-bottom: 2px solid #3b82f6;
        padding-bottom: 5px;
        color: #e2e8f0;
    }
    .report-summary {
        background-color: #1e3a5a;
        padding: 15px;
        border-radius: 8px;
        margin-bottom: 20px;
    }
    .marked-review {
        border-left: 4px solid #ff9800;
        padding-left: 10px;
        background-color: #3d3223;
    }
    </style>
    """, unsafe_allow_html=True)
    
   
    st.markdown('<h1 class="main-header">📊 Excel Mock Interview Agent</h1>', unsafe_allow_html=True)
    st.markdown("---")
    

    if not st.session_state.interview_started:
        display_interview_introduction()
        
        col1, col2, col3 = st.columns([1, 1, 1])
        with col2:
            if st.button("🚀 Start Interview", type="primary", use_container_width=True):
                st.session_state.interview_started = True
                st.rerun()
    

    elif st.session_state.interview_started and not st.session_state.interview_complete:
   
        display_agent_thinking(st.session_state.current_question + 1)
        

        progress = (st.session_state.current_question) / 5
        col1, col2 = st.columns([3, 1])
        with col1:
            st.progress(progress)
        with col2:
            st.markdown(f"<div style='text-align: center; padding: 5px; background-color: #1E1E1E; border-radius: 5px;'><strong>{st.session_state.current_question + 1}/5</strong> Questions</div>", unsafe_allow_html=True)
        

        if st.session_state.current_question == len(st.session_state.questions_answers):
            with st.spinner("Generating your next question..."):
                question_number = st.session_state.current_question + 1
                question = generate_excel_question(question_number)
                st.session_state.questions_answers.append({
                    'question': question,
                    'answer': '',
                    'score': 0,
                    'explanation': '',
                    'difficulty_level': question_number
                })
        
        current_qa = st.session_state.questions_answers[st.session_state.current_question]
  
        difficulty_levels = {
            1: {"level": "Basic", "color": "🟢"},
            2: {"level": "Intermediate-Basic", "color": "🟡"},
            3: {"level": "Intermediate", "color": "🟠"},
            4: {"level": "Advanced-Intermediate", "color": "🔴"},
            5: {"level": "Advanced", "color": "🟣"}
        }
        
        current_difficulty = difficulty_levels.get(st.session_state.current_question + 1, {"level": "Unknown", "color": "⚪"})
        
       
        st.markdown(f"""
        <div class="agent-message">
            <h3>Question {st.session_state.current_question + 1}</h3>
            <p style="font-size: 1.1rem; margin: 0.5rem 0;">{current_qa['question']}</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Answer input
        st.markdown("### Your Answer:")
        answer = st.text_area(
            "Provide your detailed answer here:",
            value="",  # Remove the default value from session state
            height=200,
            placeholder="Type your answer here... Be specific and include relevant Excel functions, steps, or formulas."
        )
        
        # Mark for review checkbox
        is_marked = st.session_state.marked_for_review[st.session_state.current_question]
        mark_for_review = st.checkbox("🚩 Mark for review", value=is_marked)
        st.session_state.marked_for_review[st.session_state.current_question] = mark_for_review
        
        # Navigation buttons with skip option
        col1, col2, col3 = st.columns([1, 1, 1])
        
        with col1:
            if st.button("📝 Save Answer", use_container_width=True):
                st.session_state.current_answer = answer
                st.session_state.questions_answers[st.session_state.current_question]['answer'] = answer
                st.success("Answer saved!")
        
        with col2:
            if st.button("⏭️ Skip Question", use_container_width=True):
                # Add to skipped questions if not already there
                if st.session_state.current_question not in st.session_state.skipped_questions:
                    st.session_state.skipped_questions.append(st.session_state.current_question)
                
                # Save any partial answer
                if answer.strip():
                    st.session_state.questions_answers[st.session_state.current_question]['answer'] = answer
                
                # Move to next question
                if st.session_state.current_question < 4:
                    st.session_state.current_question += 1
                    st.session_state.current_answer = ""
                else:
                    # All questions viewed, show submit all button
                    st.session_state.show_submit_all = True
                
                st.rerun()
        
        with col3:
            if st.button("➡️ Next Question", type="primary", use_container_width=True):
                if not answer.strip():
                    st.error("Please provide an answer before continuing.")
                else:
                    # Save the answer
                    st.session_state.questions_answers[st.session_state.current_question]['answer'] = answer
                    
                    # Remove from skipped questions if it was skipped before
                    if st.session_state.current_question in st.session_state.skipped_questions:
                        st.session_state.skipped_questions.remove(st.session_state.current_question)
                    
                    # Move to next question or show submit all button
                    st.session_state.current_question += 1
                    st.session_state.current_answer = ""
                    
                    if st.session_state.current_question >= 5:
                        # All questions answered, show submit all button
                        st.session_state.current_question = 4  # Keep at last question
                        st.session_state.show_submit_all = True
                    
                    st.rerun()
        
        # Show submit all button when all questions are answered
        if st.session_state.current_question == 4 and st.session_state.get('show_submit_all', False):
            st.markdown("---")
            st.markdown("### 🎯 Ready to Submit All Answers")
            st.markdown("You've completed all questions! Click below to submit all your answers for evaluation.")
            
            # Use the new CSS class for the submit all button
            st.markdown("<div class='submit-all-button'>", unsafe_allow_html=True)
            if st.button("📊 Submit All & Generate Report", type="primary", use_container_width=True):
                st.session_state.interview_complete = True
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)
    
    # Final Report Section
    elif st.session_state.interview_complete:
        # Evaluate all answers at once
        if not st.session_state.get('answers_evaluated', False):
            # Display evaluation in progress message
            st.markdown("""
            <div class="agent-message">
                <h2>Evaluating Your Answers...</h2>
                <p>Please wait while we analyze all your responses and prepare a comprehensive evaluation report.</p>
            </div>
            """, unsafe_allow_html=True)
            
            # Evaluate all answers
            st.session_state.questions_answers = evaluate_all_answers(st.session_state.questions_answers)
            st.session_state.answers_evaluated = True
            st.rerun()
        
        # Concluding message
        st.markdown("""
        <div class="agent-message">
            <h2>Interview Complete!</h2>
            <p>Thank you for completing the interview. We've analyzed all your responses and prepared a comprehensive evaluation report.</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Calculate total score
        total_score = sum(qa['score'] for qa in st.session_state.questions_answers)
        
        # Display individual Q&A summary
        st.markdown("### 📋 Interview Summary")
        
        # Define correct answers for each question
        correct_answers = {
            0: "To calculate the total cost, you would use the SUMPRODUCT function. In cell C7, the formula would be =SUMPRODUCT(A2:A6,B2:B6) which multiplies each item's price by its quantity and then adds all the results together.",
            1: "To create a dynamic chart that updates automatically, you would: 1) Create a named range for your data (Ctrl+T or Insert > Table), 2) Insert a chart based on this table (Insert > Charts > desired chart type), 3) The chart will automatically update when data in the table changes. You can also use OFFSET or INDEX functions with COUNTA to create dynamic ranges.",
            2: "To find the last value in column A, you can use: =LOOKUP(2,1/(A:A<>""),A:A) or =INDEX(A:A,MATCH(9.99999999999999E+307,A:A)) or =INDEX(A:A,COUNTA(A:A)). These formulas work even when the data has blank cells or is unsorted.",
            3: "To create a conditional formatting rule that highlights cells with values above average: 1) Select the range, 2) Go to Home > Conditional Formatting > New Rule, 3) Choose 'Use a formula', 4) Enter =A1>AVERAGE($A$1:$A$100) (adjust range as needed), 5) Click Format and choose highlighting style, 6) Click OK. This will highlight all cells with values above the average of the selected range.",
            4: "To create a pivot table summarizing sales by region and product: 1) Select your data range, 2) Go to Insert > PivotTable, 3) In the PivotTable Fields pane, drag 'Region' to Rows area, 'Product' to Columns area, and 'Sales' to Values area, 4) The pivot table will automatically calculate the sum of sales for each region-product combination. You can then add filters, change calculation type (e.g., to average), or add additional fields as needed."
        }
        
        for i, qa in enumerate(st.session_state.questions_answers, 1):
            score_colors = {0: "score-0", 1: "score-1", 2: "score-2"}
            score_texts = {0: "Needs Improvement", 1: "Good Understanding", 2: "Excellent Response"}
            
            # Get difficulty level for display
            difficulty_levels = {
                1: {"level": "Basic", "color": "🟢"},
                2: {"level": "Intermediate-Basic", "color": "🟡"},
                3: {"level": "Intermediate", "color": "🟠"},
                4: {"level": "Advanced-Intermediate", "color": "🔴"},
                5: {"level": "Advanced", "color": "🟣"}
            }
            
            difficulty_info = difficulty_levels.get(i, {"level": "Unknown", "color": "⚪"})
            
            # Check if question was marked for review
            marked_class = "marked-review" if st.session_state.marked_for_review[i-1] else ""
            
            with st.expander(f"Question {i} - {difficulty_info['level']} {difficulty_info['color']} - Score: {qa['score']}/2"):
                st.markdown(f"<div class='question-item {marked_class}'>", unsafe_allow_html=True)
                st.markdown(f"<div class='question-text'>Question {i}: {qa['question']}</div>", unsafe_allow_html=True)
                st.markdown(f"<div class='user-answer'>Your Answer: {qa['answer']}</div>", unsafe_allow_html=True)
                st.markdown(f"<div class='correct-answer'>Correct Answer: {correct_answers[i-1]}</div>", unsafe_allow_html=True)
                st.markdown(f"<div>Evaluation: <span class='score-indicator {score_colors[qa['score']]}'>{qa['score']}/2 - {score_texts[qa['score']]}</span></div>", unsafe_allow_html=True)
                st.markdown(f"<div>Feedback: {qa['explanation']}</div>", unsafe_allow_html=True)
                st.markdown("</div>", unsafe_allow_html=True)
        
        # Generate and display final report
        st.markdown("### 📊 Comprehensive Evaluation Report")
        
        with st.spinner("Generating comprehensive evaluation report..."):
            final_report = generate_final_report(st.session_state.questions_answers)
        
        # Calculate percentage score
        percentage_score = (total_score / 10) * 100
        
        # Display summary score in a nice format
        st.markdown(f"""
        <div class="report-summary">
            <h3>Overall Score: {total_score}/10 ({percentage_score:.1f}%)</h3>
        </div>
        """, unsafe_allow_html=True)
        
        # Determine performance level
        if percentage_score >= 90:
            performance_level = "🌟 Expert"
            performance_class = "score-2"
        elif percentage_score >= 75:
            performance_level = "✨ Advanced"
            performance_class = "score-2"
        elif percentage_score >= 60:
            performance_level = "🔵 Intermediate"
            performance_class = "score-1"
        elif percentage_score >= 40:
            performance_level = "🟡 Basic"
            performance_class = "score-1"
        else:
            performance_level = "🔴 Beginner"
            performance_class = "score-0"
        
        # Parse the final report to apply our custom styling
        # Process the final report without fancy formatting
        sections = final_report.split('**')
        formatted_report = ""
        
        for i, section in enumerate(sections):
            if i == 0:  # Skip the first empty section
                continue
                
            if i % 2 == 1:  # Section titles
                section_title = section.strip(':*')
                formatted_report += f"### {section_title}\n\n"
            else:  # Section content
                content = section.strip()
                formatted_report += f"{content}\n\n"
        
        st.markdown(f"""
        ### Final Assessment
        
        **Overall Score: {total_score}/10 ({percentage_score:.0f}%)**
        
        **Performance Level: {performance_level}**
        
        {formatted_report}
        """)
        
        # Conclusion
        st.markdown("""
        ### Interview Complete
        
        This concludes the Excel skills assessment. We hope this interview provided valuable insights into your Excel capabilities and areas for improvement. Feel free to take another interview anytime to track your progress!
        """)
        
        # Restart option
        st.markdown("---")
        col1, col2, col3 = st.columns([1, 1, 1])
        with col2:
            if st.button("🔄 Take Another Interview", type="primary", use_container_width=True):
                # Reset session state
                for key in ['interview_started', 'current_question', 'questions_answers', 'current_answer', 'interview_complete', 'answers_evaluated', 'show_submit_all']:
                    if key in st.session_state:
                        del st.session_state[key]
                st.rerun()

if __name__ == "__main__":
    main()
