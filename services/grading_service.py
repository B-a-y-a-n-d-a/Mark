import pandas as pd
import spacy
from scispacy.abbreviation import AbbreviationDetector

# Load the NLP model and add abbreviation detector once (global scope)
nlp = spacy.load("en_core_web_sm")
nlp.add_pipe("abbreviation_detector")

def resolve_abbreviation(text):
    """Expand abbreviations in text using NLP"""
    doc = nlp(text)
    mapping = {}
    for abrv in doc._.abbreviations:
        mapping[str(abrv)] = str(abrv._.long_form)
    replaced_text = text
    for abrv, long_form in mapping.items():
        replaced_text = replaced_text.replace(abrv, long_form)
    return replaced_text

def grade_student(student_file, memo_file, output_file="Marking.xlsx", threshold=0.6):
    # Load Excel files
    students = pd.read_excel(student_file, sheet_name="Sheet1")
    memo = pd.read_excel(memo_file, sheet_name="Sheet1")

    # Merge on QuestionNumber
    compare = pd.merge(students, memo, on="QuestionNumber", suffixes=('_Student', '_Memo'))
    
    # Helper functions for marking logic
    def contains_named_entity_match(answer_student, answer_memo):
        """Check if answers share common entities (case-insensitive)"""
        student_entities = {ent.text.lower() for ent in nlp(answer_student).ents}
        memo_entities = {ent.text.lower() for ent in nlp(answer_memo).ents}
        return bool(student_entities & memo_entities)

    def spacy_similarity(answer_student, answer_memo):
        """Calculate semantic similarity between answers"""
        return nlp(answer_student).similarity(nlp(answer_memo))
    
    def combined_marking(student_ans, memo_ans):
        """Apply combined NLP evaluation techniques"""
        # Resolve abbreviations and normalize case
        resolved_student = resolve_abbreviation(student_ans).strip().lower()
        resolved_memo = resolve_abbreviation(memo_ans).strip().lower()
        
        if contains_named_entity_match(resolved_student, resolved_memo):
            return 2
        return 2 if spacy_similarity(resolved_student, resolved_memo) >= threshold else 0
    
    def smart_marking(row):
        """Select appropriate marking strategy based on answer length"""
        # Handle missing values and convert to string
        student_ans = str(row['Answers_Student']).strip() if pd.notna(row['Answers_Student']) else ""
        memo_ans = str(row['Answers_Memo']).strip() if pd.notna(row['Answers_Memo']) else ""
        
        # Simple comparison for short answers
        if len(student_ans) <= 7:
            return 2 if student_ans.lower() == memo_ans.lower() else 0
        # NLP evaluation for longer answers
        return combined_marking(student_ans, memo_ans)
    
    # Apply marking logic
    compare['Mark'] = compare.apply(smart_marking, axis=1)
    
    # Calculate and add total marks
    total_marks = compare['Mark'].sum()
    compare.loc['Total'] = [''] * len(compare.columns)  # Empty row
    compare.loc['Total', 'Mark'] = total_marks

    # Export result
    compare.to_excel(output_file, index=False)
    return output_file