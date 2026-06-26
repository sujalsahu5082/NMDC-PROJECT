import os
import re
import json
import urllib.request
import urllib.error

def load_dotenv():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    env_path = os.path.join(base_dir, '..', '.env')
    if os.path.exists(env_path):
        with open(env_path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#') and '=' in line:
                    key, val = line.split('=', 1)
                    val = val.strip().strip("'\"")
                    os.environ[key.strip()] = val

# Load configurations at import time
load_dotenv()

class AIChatAgent:
    def __init__(self):
        pass

    def handle_query(self, query, data_summary, api_key=None):
        """
        Processes a user query based on the active dashboard summary.
        If api_key is provided or exists in the environment/dot-env, uses the live Groq API.
        Otherwise, falls back to the local analytical NLP engine.
        """
        query_clean = query.strip().lower()
        api_key = api_key or os.environ.get("GROQ_API_KEY")

        if api_key:
            return self._call_groq_api(query, data_summary, api_key)
        else:
            return self._handle_offline_nlp(query_clean, data_summary)

    def _call_groq_api(self, query, data_summary, api_key):
        """
        Calls the Groq API (Llama 3.3 70b) directly using Python's standard library.
        """
        url = "https://api.groq.com/openai/v1/chat/completions"
        
        # Build prompt
        context = {
            "dashboard_metrics": {
                "total_employees_under_filters": data_summary.get("total", 0),
                "kpis": data_summary.get("kpis", {}),
                "top_designations": data_summary.get("designations", [])[:5],
                "gender_breakdown": data_summary.get("gender", []),
                "deposit_distribution": data_summary.get("deposit_dist", []),
                "skills_summary": data_summary.get("skills", [])[:7],
                "active_filters": {
                    "sidebar_counts": data_summary.get("sidebar", {})
                }
            }
        }
        
        system_instruction = (
            "You are NMDC HR Analytics Intelligence Chatbot. You assist HR officers with analytical insights about employee rosters.\n"
            "You were developed by Sujal Sahu. If a user asks who developed you, who your developer is, or who created you, always respond that you were developed by Sujal Sahu.\n"
            "Here is the CURRENT FILTERED state of the dashboard data in JSON format:\n"
            f"{json.dumps(context, indent=2)}\n\n"
            "Instructions:\n"
            "1. Answer the user's question accurately using only or primarily this data.\n"
            "2. Keep your answers concise, clear, and professional.\n"
            "3. Use Markdown tables, lists, and bold text to present numbers cleanly.\n"
            "4. If the user asks general HR questions, keep it relevant to NMDC or explain that your focus is on the uploaded roster data.\n"
        )
        
        payload = {
            "model": "llama-3.3-70b-versatile",
            "messages": [
                {
                    "role": "system",
                    "content": system_instruction
                },
                {
                    "role": "user",
                    "content": query
                }
            ],
            "temperature": 0.2
        }
        
        try:
            req = urllib.request.Request(
                url,
                data=json.dumps(payload).encode('utf-8'),
                headers={
                    'Content-Type': 'application/json',
                    'Authorization': f'Bearer {api_key}',
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
                },
                method='POST'
            )
            with urllib.request.urlopen(req, timeout=10) as response:
                res_data = json.loads(response.read().decode('utf-8'))
                
                # Extract text response from OpenAI compatible format
                choices = res_data.get("choices", [])
                if choices:
                    message = choices[0].get("message", {})
                    content = message.get("content", "")
                    if content:
                        return content
                return "Received empty response from the Groq AI service."
                
        except urllib.error.HTTPError as e:
            err_msg = e.read().decode('utf-8')
            try:
                err_json = json.loads(err_msg)
                err_detail = err_json.get('error', {}).get('message', 'Unknown API error.')
            except:
                err_detail = e.reason
            
            offline_res = self._handle_offline_nlp(query.strip().lower(), data_summary)
            return (
                f"⚠️ **Groq API Error (HTTP {e.code}):** {err_detail}\n\n"
                f"*Falling back to local offline analysis:*\n\n"
                f"{offline_res}"
            )
        except Exception as e:
            offline_res = self._handle_offline_nlp(query.strip().lower(), data_summary)
            return (
                f"⚠️ **Connection Error:** Unable to reach Groq AI services ({str(e)}).\n\n"
                f"*Falling back to local offline analysis:*\n\n"
                f"{offline_res}"
            )

    def _handle_offline_nlp(self, query, data_summary):
        """
        A rule-based local analytical NLP responder that answers questions about the dashboard.
        """
        kpis = data_summary.get("kpis", {})
        total = data_summary.get("total", 0)
        
        # Helper variables
        genders = {item['label']: item['count'] for item in data_summary.get("gender", [])}
        deposits = {item['label']: item['count'] for item in data_summary.get("deposit_dist", [])}
        skills = data_summary.get("skills", [])
        desigs = data_summary.get("designations", [])
        
        # 1. Greetings
        if any(w in query for w in ["hello", "hi", "hey", "greetings", "who are you"]):
            return (
                "👋 **Welcome to NMDC HR Analytics Intelligence Chatbot!**\n\n"
                "I can analyze and summarize the active dashboard data for you.\n"
                "Ask me things like:\n"
                "- *\"Summarize the current workforce\"*\n"
                "- *\"What is the gender ratio?\"*\n"
                "- *\"Show top skills\"*\n"
                "- *\"Break down employees by deposit\"*\n\n"
                "*(Note: The AI model is currently offline. To enable live conversational AI, configure a valid Groq API Key (GROQ_API_KEY) in your server's .env file.)*"
            )
        
        # 1.5 Developer check
        if any(w in query for w in ["developer", "creator", "developed", "created", "who made", "who built", "sujal", "sahu"]):
            return "🤖 **NMDC HR Analytics Assistant** was developed by **Sujal Sahu**."

        # 2. General Summary / Overview
        if any(w in query for w in ["summary", "summarise", "overview", "what is this", "tell me about"]):
            # Find largest categories
            skills_txt = ", ".join([f"{s['label']} ({s['count']})" for s in skills[:3]]) if skills else "None"
            deps_txt = ", ".join([f"Deposit {k} ({v})" for k, v in deposits.items()]) if deposits else "None"
            
            return (
                f"### 📊 Roster Analysis Summary\n"
                f"Here is a summary of the active workforce under current filters:\n\n"
                f"- **Total Employees**: **{total}** across **{kpis.get('departments', 0)}** department(s).\n"
                f"- **Roles & Titles**: Found **{kpis.get('designations', 0)}** unique designations across **{kpis.get('grades', 0)}** grades.\n"
                f"- **Deposit Distribution**: {deps_txt}.\n"
                f"- **Gender Diversity**: **{kpis.get('female_pct', 0)}%** female representation "
                f"({genders.get('Female', 0)} female vs. {genders.get('Male', 0)} male employees).\n"
                f"- **Key Skills**: {skills_txt}.\n"
            )

        # 3. Gender specific
        if any(w in query for w in ["gender", "female", "male", "women", "men", "sex", "diversity"]):
            fem_cnt = genders.get('Female', 0)
            male_cnt = genders.get('Male', 0)
            fem_pct = kpis.get('female_pct', 0)
            male_pct = 100 - fem_pct if total > 0 else 0
            return (
                f"### 🚺 Gender Diversity Breakdown\n"
                f"- **Total Roster Count**: {total}\n"
                f"- **Female Employees**: **{fem_cnt}** ({fem_pct}%)\n"
                f"- **Male Employees**: **{male_cnt}** ({male_pct}%)\n\n"
                "This indicates the workforce is predominantly "
                f"{'male' if male_cnt >= fem_cnt else 'female'} in the filtered segment."
            )

        # 4. Deposit specific
        if any(w in query for w in ["deposit", "mines", "11b", "11c", "14"]):
            rows = ""
            for dep, count in deposits.items():
                pct = round(count / total * 100) if total else 0
                rows += f"| Deposit {dep} | **{count}** | {pct}% |\n"
            
            if not rows:
                rows = "| No deposits mapped | 0 | 0% |\n"

            return (
                f"### 📍 Deposit Distribution\n"
                f"Active employees distribution across NMDC deposits:\n\n"
                f"| Deposit Area | Employee Count | Percentage |\n"
                f"| :--- | :--- | :--- |\n"
                f"{rows}"
            )

        # 5. Skills specific
        if any(w in query for w in ["skill", "expertise", "talents", "specialty"]):
            if not skills:
                return "No skill mappings have been loaded or parsed for the current selection."
            
            rows = ""
            for idx, s in enumerate(skills[:8]):
                rows += f"{idx+1}. **{s['label']}** — {s['count']} employees\n"
            
            return (
                f"### 🛠️ Top Expertise & Skills\n"
                f"Here are the top skillsets identified in the filtered group:\n\n"
                f"{rows}"
            )

        # 6. Designation specific
        if any(w in query for w in ["designation", "title", "role", "job"]):
            if not desigs:
                return "No designation data available under the current filters."
            
            rows = ""
            for idx, d in enumerate(desigs[:5]):
                rows += f"- **{d['label']}**: {d['count']} employees\n"
            
            return (
                f"### 💼 Top Designation Distribution\n"
                f"The most common roles in this segment are:\n\n"
                f"{rows}"
            )

        # 7. Department specific
        if any(w in query for w in ["department", "dept", "sections", "production"]):
            sidebar = data_summary.get("sidebar", {})
            return (
                f"### 🏢 Department Overview\n"
                f"- **Active Departments**: {kpis.get('departments', 0)}\n"
                f"- **Production Division**: {sidebar.get('prod_all', 0)} employees\n"
                f"- **Non-Production Division**: {sidebar.get('np_all', 0)} employees\n"
                f"- **Others/Support**: {sidebar.get('oth_all', 0)} employees\n"
            )

        # 8. Help / Fallback fallback
        return (
            "🤖 **Offline Chatbot Assistant**\n\n"
            f"I have analyzed the current **{total}** active records.\n"
            "I couldn't match your request to a specific offline command. Try asking:\n"
            "- *\"What is the department breakdown?\"*\n"
            "- *\"Summarize the data\"*\n"
            "- *\"Show me top skills\"*\n"
            "- *\"How many female employees do we have?\"*\n\n"
            "**Want full Conversational AI?** Add a valid Groq API Key (`GROQ_API_KEY`) in your server's `.env` file to let the AI answer arbitrary inquiries."
        )
