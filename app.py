import os
import re
import tensorflow as tf
import spacy

# Initialize SpaCy model
nlp = spacy.load("en_core_web_sm")

# Initialize T5 model
model = tf.saved_model.load('t5-base')

def extract_functions_from_bas(bas_path):
    with open(bas_path, 'r') as file:
        content = file.read()
        functions = re.findall(r'Sub\s+(\w+)\(', content)
        return functions, content

def generate_function_description(func_name, function_code):
    # Remove excess whitespace and extract relevant parts of the function
    function_code = function_code.strip()
    
    # Summarize the first two sentences using SpaCy
    doc = nlp(function_code)
    first_two_sentences = " ".join(sent.text for sent in list(doc.sents)[:2])
    
    # Use T5 for summarization
    input_text = f"summarize: {first_two_sentences}"
    input_ids = tf.constant([tokenizer.encode(input_text, max_length=512, truncation=True)])
    
    output = model(input_ids)
    summary = tokenizer.decode(output.numpy()[0], skip_special_tokens=True)
    
    return f"The function {func_name} performs {summary}"

def process_vba_functions(bas_path):
    vba_functions, function_code = extract_functions_from_bas(bas_path)
    print(f"Functions: {vba_functions}")
    for func_name in vba_functions:
        function_description = generate_function_description(func_name, function_code)
        print(function_description)

# Example usage
if __name__ == "__main__":
    bas_path = r"D:\SG Hackathon\macro_code.bas"
    process_vba_functions(bas_path)
