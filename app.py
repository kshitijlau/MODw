import streamlit as st
import pandas as pd
import io
import re
import json
from openai import AzureOpenAI # Use the OpenAI library for Azure

# --- Helper Functions ---

def get_sample_scores_df():
    """Creates a sample DataFrame for the scores file."""
    data = {
        'Person': ['Indicator Text', 'EO1'],
        'Adaptability': ['Adaptability', 4],
        'Adaptability 1': ["Effectively navigates and leads teams through changes, minimizing disruption and maintaining morale.", 4],
        'Adaptability 2': ["Quickly learns from experiences and applies insights to new situations, demonstrating a commitment to continuous improvement.", 4],
        'Adaptability 3': ["Welcomes diverse perspectives and ideas, encouraging creative problem-solving and innovation.", 4],
        'Adaptability 4': ["Displays personal resilience, remains calm and effective in times of crisis and ambiguity.", 4],
        'Capability Development': ['Capability Development', 3],
        'Capability Development 1': ["Identifies current skills and competencies within the team and assesses gaps relative to future needs, informing targeted development initiatives.", 2.5],
        'Capability Development 2': ["Engages in coaching to develop team members' skills, providing guidance and support.", 3.5],
        'Capability Development 3': ["Delegates responsibilities effectively encouraging team members to take ownership of their work.", 3],
        'Capability Development 4': ["Proactively identifies and nurtures high-potential team members, ensuring that the organization has the necessary talent to meet current and future challenges.", 3],
        'Decision Making and Takes Accountability': ['Decision Making and Takes Accountability', 4.8],
        'Decision Making and Takes Accountability 1': ["Show the ability to act assertively and take independent and tough decisions even when they are unpopular.", 4.5],
        'Decision Making and Takes Accountability 2': ["Displays confidence and credibility in decision-making, skilfully articulating decisions to garner support and alignment from others.", 5],
        'Decision Making and Takes Accountability 3': ["Identifies potential risks associated with tactical decisions and evaluates their implications on success of the overall goals.", 4.5],
        'Decision Making and Takes Accountability 4': ["Utilizes critical thinking to assess options and make informed decisions that align with objectives and values.", 5],
        'Effective Communication and Influence': ['Effective Communication and Influence', 3.5],
        'Effective Communication and Influence 1': ["Clearly articulates ideas and information in ensuring understanding.", 4],
        'Effective Communication and Influence 2': ["Seeks common ground and influences others towards win-win outcomes, facilitating agreement between different parties.", 4],
        'Effective Communication and Influence 3': ["Demonstrates strong listening skills, ensuring that team members feel heard and understood.", 2.5],
        'Effective Communication and Influence 4': ["Adjusts communication style and approach based on the audience and context, ensuring effective engagement with diverse groups.", 3.5],
        'Initiative': ['Initiative', 3.8],
        'Initiative 1': ["Takes the initiative to identify and pursue opportunities, demonstrating a willingness to act without being prompted.", 4],
        'Initiative 2': ["Sets ambitious objectives and consistently seeks ways to exceed expectations, demonstrating a strong commitment to achieving results.", 4],
        'Initiative 3': ["Displays grit in the achievement of challenging goals, pushing boundaries for self and others performance.", 3.5],
        'Initiative 4': ["Consistently takes action beyond immediate responsibilities to achieve goals.", 3.5],
        'Inspirational Leadership': ['Inspirational Leadership', 3.4],
        'Inspirational Leadership 1': ["Develops a sense of common vision and purpose in one's team that drives activity and creates motivation to achieve overall goals.", 4],
        'Inspirational Leadership 2': ["Collaborates and works with others effectively, demonstrating the ability to judge what is the most appropriate leadership style (e.g. directive, collaborative, etc.)", 3.5],
        'Inspirational Leadership 3': ["Demonstrates awareness of one’s own emotions and those of others, is aware of his/her impact on others and uses this understanding to inspire others.", 3],
        'Inspirational Leadership 4': ["Recognizes the individual styles of each team member and proactively manages them in ways that draw out their best contributions.", 3],
        'Strategic Thinking': ['Strategic Thinking', 4],
        'Strategic Thinking 1': ["Monitors and predicts key trends in the industry to inform the future direction of the organization.", 4],
        'Strategic Thinking 2': ["Identifies and assesses potential disruptors and develops strategies to proactively navigate them.", 4],
        'Strategic Thinking 3': ["Proactively identifies new opportunities that align with organizational goals and capabilities.", 4],
        'Strategic Thinking 4': ["Translates complex strategic organizational goals into meaningful actions across teams and functions.", 4],
        'Systematic Analysis and Planning': ['Systematic Analysis and Planning', 2.8],
        'Systematic Analysis and Planning 1': ["Delivers high-quality results consistently, demonstrating effective project management skills.", 3],
        'Systematic Analysis and Planning 2': ["Creates detailed action plans that outline the steps, resources, and timelines required to achieve strategic objectives, ensuring effective execution and accountability.", 3],
        'Systematic Analysis and Planning 3': ["Effectively allocates resources (time, personnel, budget) to optimize project outcomes and align with strategic priorities.", 2.5],
        'Systematic Analysis and Planning 4': ["Establishes metrics and benchmarks to evaluate progress and effectiveness of plans, making adjustments as necessary to achieve desired results.", 2.5]
    }
    return pd.DataFrame(data)

def get_sample_comments_df():
    """Creates a sample DataFrame for the comments file."""
    data = {
        'Person Code': ['EO1', 'EO1', 'E32', 'E32'],
        'Comments': [
            'He needs to be more vocal in leadership meetings.',
            'His project planning documents are very detailed and helpful, but he sometimes misses the bigger picture on resource allocation.',
            'A bit quiet, but very reliable once a task is assigned.',
            'Would like to see him present his ideas with more confidence to senior stakeholders.'
        ]
    }
    return pd.DataFrame(data)


def df_to_excel_bytes(df):
    """Converts a DataFrame to an in-memory Excel file (bytes)."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    output.seek(0)
    return output.getvalue()


def get_score_summary_prompt():
    """Returns the prompt for generating the main summary from scores."""
    return """
**## Persona**
You are an expert talent management analyst and a master writer. Your style is formal, professional, objective, and constructive. You synthesize quantitative performance data into a rich, qualitative, behavioral-focused narrative.

**## Core Objective**
Generate a sophisticated, **exactly two-paragraph** English performance summary based on scores from 8 leadership competencies.

**## Input Data Profile**
You will receive a data set for one individual containing: 8 Competency Names and their average scores, plus 4 Indicator Scores and Texts for each competency.

**## CRITICAL WRITING RULES**
1.  **DIRECT ADDRESS:** You MUST address the candidate directly using "you" and "your". Do not use third-person "he/his" or mention the candidate's name/code (e.g., "E01") in the summary body. Start sentences like "Your clear strength lies in your ability to..."
2.  **NO COMPETENCY NAMES:** You MUST NOT use the literal competency names (e.g., 'Adaptability', 'Strategic Thinking') in the final summary.
3.  **USE VERB PHRASES:** Instead of names, you MUST describe the competency as a behavior or skill using a verb phrase.
    * **Example INSTEAD OF:** "You are strong in Strategic Thinking."
    * **Example DO THIS:** "You demonstrate a strong ability to think strategically and connect long-term goals to daily actions."
4.  **NO BOLDING:** Do not use markdown for bolding (`**`) or any other special formatting in the output.
5.  **STRICT 2-PARAGRAPH STRUCTURE:** The output must always have exactly two paragraphs before the optional comments are added.

**## Core Logic & Execution Flow**
1.  **Analyze and Group:** Mentally sort the 8 competencies by their average scores, from highest to lowest.
2.  **Mandatory Opening:** The first paragraph MUST begin with this exact text: "Your participation in the assessment center provided insight into how you demonstrate the leadership competencies in action. The feedback below highlights observed strengths and opportunities for development to support your continued growth."
3.  **Paragraph 1 (Strengths Narrative):**
    * This paragraph should cover your strongest areas, typically the top 4 competencies.
    * Weave these strengths into a cohesive narrative. Start with the most prominent strength and transition smoothly to others.
    * For each strength, synthesize the highest-scoring indicator texts into a rich, descriptive sentence that explains *how* you demonstrate that skill.
4.  **Paragraph 2 (Development Narrative):**
    * This paragraph should cover the areas with the most opportunity for growth, typically the bottom 4 competencies.
    * Frame these points constructively. For competencies that are still positive but lower-scoring, you can introduce them with phrases like "To further enhance your effectiveness...".
    * For clear development areas, be direct but professional.
    * For each point, synthesize the lowest-scoring indicator texts to explain the development opportunity. For example, if 'listening skills' is a low-scoring indicator within a communication competency, the summary should mention the need to "enhance your listening skills to ensure all team members feel fully heard."

**## Final Output Constraints**
* **Word Count:** Maximum 400 words (excluding the mandatory opening).
* **Source Fidelity:** Base all statements *strictly* on the indicator language provided.
* **No Scores:** The summary must NEVER mention specific numerical scores or averages.

---
**## TASK: GENERATE SCORE-BASED SUMMARY FOR THE FOLLOWING PERSON**
"""

def get_comment_summary_prompt():
    """Returns the new, specialized prompt for summarizing qualitative comments."""
    return """
**## Persona**
You are a discerning talent management analyst, skilled at synthesizing raw, unstructured feedback into a concise and professional **English** summary. Your focus is purely on constructive, developmental themes.

**## Core Objective**
Analyze a list of raw comments for an individual and generate a single, final summary paragraph in English. This paragraph should be no more than 50 words.

**## Input Data Profile**
1.  **The Main Report:** The already-written, score-based summary.
2.  **Raw Comments:** A list of verbatim comments from colleagues.

**## Core Logic & Execution Flow**
1.  **Filter Comments:** First, you MUST filter the raw comments based on these rules:
    * **IGNORE:** Offensive, irrelevant, purely personal, or overly judgmental comments.
    * **FOCUS ON:** Developmental aspects, constructive criticism, and actionable feedback.
2.  **Check for Contradictions:** **This is the most important rule.** Compare the themes in the filtered comments with the main report provided. If a comment's theme directly contradicts a "Clear Strength" identified in the main report, you MUST ignore that comment. The main report is the primary source of truth.
3.  **Synthesize Themes:** From the remaining, non-contradictory comments, identify 1-2 key developmental themes. If the comments are varied, select the most impactful points.
4.  **Draft the Summary:** Write a single paragraph that summarizes these themes.
    * **Introduction:** Start with a phrase like "Additionally, feedback suggests..." or "Further feedback indicates...".
    * **Body:** Concisely state the key themes. Rephrase any judgmental language into professional, developmental terms (e.g., "He is too quiet" becomes "you would benefit from increasing your visibility in senior forums.").
5.  **Final Polish:** Ensure the paragraph flows naturally when appended to the main report.

**## Writing Standards & Constraints**
* **Word Count:** Maximum 50 words.
* **Tone:** Professional, constructive, and forward-looking.
* **Consistency:** The summary MUST NOT contradict the main report.

---
**## TASK: ANALYZE THE FOLLOWING COMMENTS AND GENERATE A 50-WORD SUMMARY PARAGRAPH TO APPEND TO THE MAIN REPORT PROVIDED.**
"""

def get_translation_prompt():
    """Returns a new, dedicated prompt for translating text to Arabic."""
    return """
**## Persona**
You are an expert translator specializing in professional HR and talent management content.

**## Core Objective**
Translate the provided English text into formal, professional Arabic (`Lughat al-Fusha`).

**## Core Logic & Execution Flow**
1.  Read the English text provided after the '---' delimiter.
2.  Translate it into Arabic, ensuring the translation is not literal but captures the professional nuance, tone, and intent of the original text.
3.  The opening sentence about "participation in the assessment center" must be particularly formal and well-written.
4.  The translation must adhere to the same narrative style as the English, describing behaviors with verb phrases rather than using direct competency names.

**## Writing Standards & Constraints**
* **Language:** Formal, written Arabic.
* **Tone:** Professional, respectful, and constructive.
* **Accuracy:** Preserve the original meaning perfectly.

---
"""

def call_azure_openai(prompt):
    """A single, reusable function to call the Azure OpenAI API."""
    try:
        azure_endpoint = st.secrets["azure_endpoint"]
        api_key = st.secrets["azure_api_key"]
        deployment_name = st.secrets["azure_deployment_name"]
        api_version = "2024-02-01"

        client = AzureOpenAI(
            azure_endpoint=azure_endpoint,
            api_key=api_key,
            api_version=api_version,
        )

        message_text = [{"role": "user", "content": prompt}]

        completion = client.chat.completions.create(
            model=deployment_name,
            messages=message_text,
            temperature=0.7,
            max_tokens=1500,
            top_p=0.95,
            frequency_penalty=0,
            presence_penalty=0,
            stop=None
        )
        return completion.choices[0].message.content.strip()

    except KeyError as e:
        st.error(f"Missing Secret: Could not find '{e}'. Please check your Streamlit Cloud secrets.")
        return f"Error: Missing configuration for '{e}'."
    except Exception as e:
        st.error(f"An error occurred while calling the OpenAI API: {e}")
        return "Error: API call failed."


def process_scores(df):
    """Processes the scores dataframe to generate initial summaries."""
    results = []
    indicator_definitions = df.iloc[0]
    people_data = df.iloc[1:]
    score_prompt_template = get_score_summary_prompt()
    translation_prompt_template = get_translation_prompt()

    progress_bar = st.progress(0)
    total_people = len(people_data)
    for i, (_, row) in enumerate(people_data.iterrows()):
        person_name = row.iloc[0]
        if pd.isna(person_name) or 'ERROR' in str(row.iloc[1]): continue
        
        st.write(f"Generating English summary for {person_name}...")

        person_data_prompt = f"**Person's Name:** {person_name}\n\n**Competency Data:**\n"
        for j in range(8):
            comp_col_index = 1 + (j * 5)
            if comp_col_index >= len(df.columns): break
            person_data_prompt += f"\n**- Competency: {df.columns[comp_col_index]}** (Average Score: {row[comp_col_index]})\n"
            for k in range(4):
                ind_col_index = comp_col_index + 1 + k
                if ind_col_index >= len(df.columns): break
                person_data_prompt += f"  - Indicator: '{indicator_definitions[ind_col_index]}' | Score: {row[ind_col_index]}\n"

        # --- TWO-STEP GENERATION ---
        # 1. Generate English Summary
        full_eng_prompt = score_prompt_template + person_data_prompt
        eng_summary = call_azure_openai(full_eng_prompt)

        # 2. Generate Arabic Translation
        st.write(f"Translating summary for {person_name} to Arabic...")
        full_ar_prompt = translation_prompt_template + eng_summary
        ar_summary = call_azure_openai(full_ar_prompt)
        
        results.append({"Person": person_name, "English Summary": eng_summary, "Arabic Summary": ar_summary})
        progress_bar.progress((i + 1) / total_people)
        
    return pd.DataFrame(results)

def process_comments_and_append(results_df, comments_df):
    """Processes comments and appends them to the existing summaries."""
    comment_prompt_template = get_comment_summary_prompt()
    translation_prompt_template = get_translation_prompt()
    
    progress_bar = st.progress(0)
    total_people = len(results_df)
    for i, row in results_df.iterrows():
        person_code = row['Person']
        main_eng_summary = row['English Summary']
        
        person_comments = comments_df[comments_df['Person Code'] == person_code]['Comments'].tolist()

        if person_comments:
            st.write(f"Summarizing comments for {person_code}...")
            comment_data_prompt = f"**Main Report:**\n{main_eng_summary}\n\n**Raw Comments to Summarize:**\n- {'\n- '.join(person_comments)}"
            
            # --- TWO-STEP GENERATION FOR COMMENTS ---
            # 1. Generate English Comment Summary
            full_eng_prompt = comment_prompt_template + comment_data_prompt
            eng_comment_summary = call_azure_openai(full_eng_prompt)

            # 2. Translate English Comment Summary to Arabic
            st.write(f"Translating comments for {person_code}...")
            full_ar_prompt = translation_prompt_template + eng_comment_summary
            ar_comment_summary = call_azure_openai(full_ar_prompt)

            results_df.at[i, 'English Summary'] += f"\n\n{eng_comment_summary}"
            results_df.at[i, 'Arabic Summary'] += f"\n\n{ar_comment_summary}"
        
        progress_bar.progress((i + 1) / total_people)
            
    return results_df

# --- Streamlit App UI ---

st.set_page_config(layout="wide")
st.title("📄 Integrated Performance Summary Generator (Azure OpenAI)")

with st.expander("Secrets Debug Information (Temporary)"):
    st.write("This section helps diagnose issues with secrets on Streamlit Cloud.")
    if all(k in st.secrets for k in ["azure_endpoint", "azure_api_key", "azure_deployment_name"]):
        st.success("All required secrets (azure_endpoint, azure_api_key, azure_deployment_name) are loaded successfully!")
        st.write("Endpoint:", st.secrets["azure_endpoint"])
    else:
        st.error("One or more required secrets are missing. Please check your secrets configuration on the Streamlit Community Cloud settings page.")
        st.write("Keys found in secrets:", list(st.secrets.keys()))
st.markdown("---")


st.info("""
    **First-Time Setup:** This application requires an Azure OpenAI API key. On Streamlit Cloud, go to your app's "Settings" > "Secrets" and add the following keys:
    ```toml
    azure_endpoint = "YOUR_AZURE_OPENAI_ENDPOINT"
    azure_api_key = "YOUR_AZURE_OPENAI_API_KEY"
    azure_deployment_name = "YOUR_GPT4o_DEPLOYMENT_NAME"
    ```
""")

st.markdown("### 1. Upload Quantitative Scores File")
with st.expander("Show Score File Instructions"):
    st.write("Upload an Excel file with competency scores. The first row should be headers, the second row must contain indicator definitions, and subsequent rows should have person IDs and scores.")
    sample_scores_df = get_sample_scores_df()
    st.download_button(
        label="📥 Download Scores Template",
        data=df_to_excel_bytes(sample_scores_df),
        file_name="sample_scores_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

uploaded_scores_file = st.file_uploader("Choose a scores file", type="xlsx", key="scores_uploader")

if uploaded_scores_file:
    try:
        scores_df = pd.read_excel(uploaded_scores_file, engine='openpyxl')
        if st.button("Generate Summaries from Scores", key="generate_scores"):
            with st.spinner("Analyzing scores and generating summaries via Azure OpenAI... This may take a moment."):
                results_df = process_scores(scores_df)
                st.session_state['results_df'] = results_df
                st.success("Score-based summaries generated successfully!")
    except Exception as e:
        st.error(f"Error processing scores file: {e}")

if 'results_df' in st.session_state:
    st.markdown("---")
    st.markdown("### 2. Score-Based Summaries (Preview)")
    st.dataframe(st.session_state['results_df'].head())

    st.markdown("---")
    st.markdown("### 3. (Optional) Upload Qualitative Comments File")
    with st.expander("Show Comments File Instructions"):
        st.write("To enrich the report, upload an Excel file with raw comments. It should have two columns: 'Person Code' and 'Comments'.")
        sample_comments_df = get_sample_comments_df()
        st.download_button(
            label="📥 Download Comments Template",
            data=df_to_excel_bytes(sample_comments_df),
            file_name="sample_comments_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    uploaded_comments_file = st.file_uploader("Choose a comments file", type="xlsx", key="comments_uploader")

    if uploaded_comments_file:
        try:
            comments_df = pd.read_excel(uploaded_comments_file, engine='openpyxl')
            if st.button("Incorporate Comments into Summaries", key="generate_comments"):
                with st.spinner("Analyzing comments and updating summaries via Azure OpenAI..."):
                    current_results = st.session_state['results_df'].copy()
                    final_df = process_comments_and_append(current_results, comments_df)
                    st.session_state['final_df'] = final_df
                    st.success("Comments incorporated successfully!")
        except Exception as e:
            st.error(f"Error processing comments file: {e}")

if 'final_df' in st.session_state:
    st.markdown("---")
    st.markdown("### 4. Final Integrated Report")
    st.dataframe(st.session_state['final_df'])
    st.download_button(
        label="📥 Download Final Integrated Report",
        data=df_to_excel_bytes(st.session_state['final_df']),
        file_name="final_integrated_summaries.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
elif 'results_df' in st.session_state and 'final_df' not in st.session_state:
    st.markdown("---")
    st.markdown("### 4. Download Score-Based Report")
    st.download_button(
        label="📥 Download Score-Based Report",
        data=df_to_excel_bytes(st.session_state['results_df']),
        file_name="score_based_summaries.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
