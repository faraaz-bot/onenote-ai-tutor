import os
import requests
from google import genai

class OneNoteAITutorPipeline:
    def __init__(self, access_token, gemini_api_key):
        self.access_token = access_token
        self.headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        self.gemini_client = genai.Client(api_key=gemini_api_key)

    def analyze_course_context(self, section_id):
        url = f"https://graph.microsoft.com/v1.0/me/onenote/sections/{section_id}/pages"
        res = requests.get(url, headers=self.headers)
        if res.status_code != 200:
            return "Standard course curriculum."
            
        exam_texts = []
        for page in res.json().get("value", []):
            title = page.get("title", "").lower()
            if any(k in title for k in ["midterm", "final", "exam", "quiz", "test"]):
                p_id = page["id"]
                c_res = requests.get(f"https://graph.microsoft.com/v1.0/me/onenote/pages/{p_id}/content", headers=self.headers)
                if c_res.status_code == 200:
                    exam_texts.append(c_res.text)
                    
        if exam_texts:
            prompt = f"Analyze these past exams/midterms to identify core recurring high-yield themes and weighting:\n\n" + "\n".join(exam_texts[:3])
            try:
                response = self.gemini_client.models.generate_content(
                    model='gemini-2.5-flash',
                    contents=prompt
                )
                return response.text
            except Exception as e:
                return f"Exam context extraction note: {e}"
        return "Standard course curriculum themes."

    def generate_ai_content(self, slide_text, exam_context, slide_title):
        prompt = f"""
        You are an expert AI professor and tutor. Analyze the following lecture slide content and past exam context to produce study notes.
        
        Slide Title: {slide_title}
        Slide Content / OCR:
        {slide_text}
        
        Course Midterm/Final Exam Themes:
        {exam_context}
        
        Generate output in valid HTML format (using inline CSS) structured as follows:
        1. Detailed Explanation of the concepts.
        2. Relevance section mapping this slide specifically to the course midterm/final exam themes.
        """
        
        try:
            response = self.gemini_client.models.generate_content(
                model='gemini-2.5-flash',
                contents=prompt
            )
            ai_html = response.text
        except Exception as e:
            ai_html = f"<p>AI Generation Error: {e}</p>"
            
        return f"""
        <div style="background-color: #f9f9fb; padding: 20px; border-radius: 8px; border: 1px solid #e1dfdd; margin-top: 20px;">
            <h2 style="color: #0078d4;">🤖 AI Tutor Explanation & Exam Relevance</h2>
            {ai_html}
        </div>
        """

    def generate_quiz_section(self, section_text):
        prompt = f"""
        Based on the following lecture materials, create 3 rigorous multiple-choice practice quiz questions with answers and explanations. Output in clean HTML format.
        
        Lecture Material:
        {section_text[:3000]}
        """
        try:
            response = self.gemini_client.models.generate_content(
                model='gemini-2.5-flash',
                contents=prompt
            )
            quiz_html = response.text
        except Exception as e:
            quiz_html = f"<p>Quiz Generation Error: {e}</p>"
            
        return f"""
        <div style="background-color: #f3f2f1; padding: 20px; border-radius: 8px; border: 1px solid #c8c6c4; margin-top: 30px;">
            <h2 style="color: #107c10;">📝 Lecture Practice Quiz</h2>
            {quiz_html}
        </div>
        """

    def process_lecture_section(self, section_id):
        exam_context = self.analyze_course_context(section_id)
        
        url = f"https://graph.microsoft.com/v1.0/me/onenote/sections/{section_id}/pages"
        res = requests.get(url, headers=self.headers)
        if res.status_code != 200:
            return [f"Failed to fetch pages: {res.text}"]
            
        pages = res.json().get("value", [])
        results = []
        all_slide_text = ""
        
        for page in pages:
            p_id = page["id"]
            p_title = page["title"]
            
            if "AI Tutor" in p_title or "Quiz" in p_title:
                continue
                
            c_res = requests.get(f"https://graph.microsoft.com/v1.0/me/onenote/pages/{p_id}/content", headers=self.headers)
            slide_text = c_res.text if c_res.status_code == 200 else ""
            all_slide_text += slide_text + "\n"
            
            ai_html = self.generate_ai_content(slide_text, exam_context, p_title)
            
            patch_url = f"https://graph.microsoft.com/v1.0/me/onenote/pages/{p_id}/content"
            patch_payload = [{
                "target": "body",
                "action": "append",
                "position": "after",
                "content": ai_html
            }]
            
            p_res = requests.patch(patch_url, headers=self.headers, json=patch_payload)
            if p_res.status_code in (200, 204):
                results.append(f"Added AI notes to: {p_title}")
                
        quiz_html = self.generate_quiz_section(all_slide_text)
        create_page_url = f"https://graph.microsoft.com/v1.0/me/onenote/sections/{section_id}/pages"
        quiz_page_body = f"""
        <!DOCTYPE html>
        <html>
          <head><title>Practice Quiz & Review</title></head>
          <body>{quiz_html}</body>
        </html>
        """
        requests.post(create_page_url, headers={"Authorization": self.headers["Authorization"], "Content-Type": "application/xhtml+xml"}, data=quiz_page_body.encode('utf-8'))
        
        return results
