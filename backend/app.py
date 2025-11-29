from flask import Flask, request, jsonify
from flask_cors import CORS
import openai
import os
from dotenv import load_dotenv

load_dotenv()  # Load environment variables from .env file

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

# Initialize OpenAI client
client = openai.OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

@app.route('/rewrite', methods=['POST'])
def rewrite_text():
    try:
        data = request.get_json()
        original_text = data.get('text', '')
        instruction = data.get('instruction', 'Rewrite this text in a formal academic tone')
        
        if not original_text.strip():
            return jsonify({'error': 'No text provided'}), 400
        
        # Call OpenAI API
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are a helpful writing assistant. Rewrite the user's text according to their instructions. Return the result as HTML suitable for pasting into Microsoft Word (e.g., use <p>, <strong>, <em>, <ul>, <li>). Do not include <html> or <body> tags, just the content."},
                {"role": "user", "content": f"Instruction: {instruction}\n\nText: {original_text}"}
            ]
        )
        
        rewritten_text = response.choices[0].message.content
        return jsonify({'rewritten_text': rewritten_text})
        
    except Exception as e:
        print(f"Error: {str(e)}")  # Log error to console
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
