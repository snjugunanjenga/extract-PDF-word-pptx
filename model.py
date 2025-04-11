from transformers import pipeline
from docx import Document



def generate_transcript_from_text(text, max_length=2000):

    # Initialize the GPT model for text generation (GPT-3.5 or GPT-4 would be best suited)
    model = pipeline("text-generation", model="gpt-3.5-turbo")  # You can use GPT-4 or T5 for more complex tasks

    # Ask the model to generate a detailed, informative transcript (targeting a 10-minute length)
    prompt = f"Generate a 10-minute lecture-style transcript for a deep learning engineer, explaining the following technical concepts in detail, step by step: {text}"

    # Generate transcript with higher max_length for detailed content
    transcript = model(prompt, max_length=max_length, num_return_sequences=1)

    return transcript[0]['generated_text']

# Example usage
extract_text = r'C:\Users\People\Desktop\upwork\PPT deep learning\extract_text.docx'

# Assuming 'extract_text' contains the required content for the transcript
transcript_text = generate_transcript_from_text(extract_text)

doc = Document()
transcript_path = r'C:\Users\People\Desktop\upwork\PPT deep learning\extract_text.docx'
doc.save(transcript_path)
print(f"Document saved at {transcript_path}")
