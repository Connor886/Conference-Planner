import streamlit as st
from dotenv import load_dotenv
import os
import re
import time
import json
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from langchain_core.tools import tool
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain_tavily import TavilySearch
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.enum.section import WD_ORIENT

# Load environmental vairables
load_dotenv()

# ─────────────────────────────────────────────────────────────────────────────
# Tools
# ─────────────────────────────────────────────────────────────────────────────

@tool
def tavily_search_tool(query: str) -> str:
    '''Finds the most relevant conference website URL using Tavily'''
    try:
        tool = TavilySearch(
            api_key = os.getenv('TAVILY_API_KEY'),
            max_results = 5,
            topic = 'general',
        )
        result = tool.invoke({'query': f'Find the official website for the full schedule page of the {query} pharmaceutical conference.'})
        for r in result.get('results', []):
            if 'url' in r:
                return r['url']
        return 'No conference URL found.'
    except Exception as e:
        return f'Tavily search failed: {str(e)}'

@tool
def scrape_conference_website_tool(url: str) -> str:
    '''Scrapes rendered HTML content from the given conference URL using Selenium'''
    try:
        options = Options()
        options.binary_location = '/usr/bin/google-chrome'
        options.add_argument('--headless')
        options.add_argument('--disable-gpu')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')

        driver = webdriver.Chrome(service = Service(ChromeDriverManager().install()), options = options)
        driver.get(url)
        time.sleep(5)

        for text in ['Full Schedule', 'Agenda', 'Conference', 'Program', 'Presentation', 'Schedule', 'Session', 'Sessions']:
            try:
                link = driver.find_element(By.PARTIAL_LINK_TEXT, text)
                link.click()
                time.sleep(5)
                break
            except:
                continue

        try:
            iframes = driver.find_elements(By.TAG_NAME, 'iframe')
            for iframe in iframes:
                try:
                    driver.switch_to.frame(iframe)
                    iframe_text = driver.page_source.lower()
                    if any(kw in iframe_text for kw in ['session', 'event', 'schedule']):
                        print('Switched to iframe with session content')
                        for _ in range(3):
                            driver.execute_script('window.scrollBy(0, document.body.scrollHeight);')
                            time.sleep(3)
                        break
                    driver.switch_to.default_content()
                except Exception as e:
                    print('Failed switching iframe:', e)
            print('Switched to iframe for scraping')
        except Exception as e:
            print('No iframe found or unable to switch:', e)
        html = driver.page_source
        driver.quit()
        try:
            with open('/tmp/scraped_iframe.html', 'w') as f:
                f.write(html)
        except:
            pass
        soup = BeautifulSoup(html, 'html.parser')
        candidates = []
        keywords = ['session', 'presenter', 'time', 'date', 'title', 'speaker', 'event']
        for container in soup.find_all(['div', 'section', 'ul']):
            children = container.find_all(recursive = True)
            if len(children) >= 3:
                text_samples = [child.get_text(strip = True) for child in children[:3]]
                keyword_hits = 0
                for sample in text_samples:
                    sample_str = sample if isinstance(sample, str) else str(sample)
                    if any(kw in sample_str.lower() for kw in keywords):
                        keyword_hits += 1
                if keyword_hits >= 1:
                    candidates.append((container, len(children)))
        if not candidates:
            return 'No suitable session container found automatically'
        candidates.sort(key = lambda x: x[1], reverse = True)
        best_container = candidates[0][0]
        text = best_container.get_text(separator = '\n', strip = True)
        return text[:15000]   
    except Exception as e:
        return f'Scraping failed: {e}'
    
@tool
def linkedin_search_tool(name_and_affiliation: str) -> str:
    '''Searches for the LinkedIn profile of a speaker using Tavily and extracts structured profile fields'''
    try: 
        tool = TavilySearch(
            api_key = os.getenv('TAVILY_API_KEY'),
            max_results = 3
        )
        results = tool.invoke({'query': f'{name_and_affiliation} site: linkedin.com'})
        for r in results.get('results', []):
            if 'linkedin.com/in/' in r.get('url', ''):
                snippet = r.get('content', '')
                llm_prompt = f'''
Extract the following fields from this LinkedIn search snippet:

- Professional Title
- Institution or Company
- City and State (if available)
- Short Biography (1-2 sentences max)

Snippet:
\"\"\"
{snippet}
\"\"\"

Return as JSON with keys: professional_title, institution, city_state, bio.
'''
                llm_response = llm.invoke(llm_prompt)
                parsed = safe_parse_json(llm_response.content)
                return {
                    'name': name_and_affiliation,
                    'linkedin_url': r['url'],
                    'professional_title': parsed.get('professional_title', 'N/A'),
                    'institution': parsed.get('institution', 'N/A'),
                    'city_state': parsed.get('city_state', 'N/A'),
                    'bio': parsed.get('bio', 'N/A')
                }
        search_query = name_and_affiliation.strip().replace(' ', '%20')
        search_bar_url = f'https://www.linkedin.com/search/results/people/?keywords={search_query}'
        return {
            'name': name_and_affiliation,
            'linkedin_url': 'N/A',
            'search_bar_url': search_bar_url,
            'professional_title': 'N/A',
            'institution': 'N/A',
            'city_state': 'N/A',
            'bio': 'N/A'
        }
    except Exception as e:
        return {
            'name': name_and_affiliation,
            'linkedin_url': 'N/A',
            'professional_title': 'N/A',
            'institution': 'N/A',
            'city_state': 'N/A',
            'bio': 'N/A'
        }

@tool
def table_tool(info: list[dict]):
    '''Creates a DOCX table from selected conference sessions'''
    document = Document()
    section = document.sections[-1]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width

    table = document.add_table(rows = len(info)+1, cols = 7)
    set_table_border(table)

    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Name of Presenter(s)'
    hdr_cells[1].text = 'Professional Title'
    hdr_cells[2].text = 'Institution'
    hdr_cells[3].text = 'City/State'
    hdr_cells[4].text = 'Title of Presentation'
    hdr_cells[5].text = 'Date, Time, and Location'
    hdr_cells[6].text = 'LinkedIn Profile'

    for i, dict in enumerate(info):
        row_cells = table.rows[i+1].cells
        row_cells[0].text = dict.get('names') or ''
        row_cells[1].text = dict.get('professional_titles') or ''
        row_cells[2].text = dict.get('institution') or ''
        row_cells[3].text = dict.get('city_state') or ''
        row_cells[4].text = dict.get('title') or ''
        row_cells[5].text = dict.get('dtl') or ''
        linkedin_url = dict.get('linkedin_url') or ''
        if linkedin_url and linkedin_url != 'N/A':
            add_hyperlink(row_cells[6].paragraph[0], linkedin_url, 'LinkedIn Profile')
        else:
            row_cells[6].text = 'N/A'

    document.save('table.docx')
    return 'Table successfully created!'

# ─────────────────────────────────────────────────────────────────────────────
# Helper functions
# ─────────────────────────────────────────────────────────────────────────────

def extract_presenter_names(llm_output: str) -> list[str]:
    presenter_pattern = r'\*\*Presenter(?:\(s\))?:\*\*\s*(.+)'
    matches = re.findall(presenter_pattern, llm_output)
    names = []
    for match in matches:
        split_names = re.split(r',\s*and\s*|, \s|\sand\s', match)
        for name in split_names:
            name = name.strip()
            if name and name not in names:
                names.append(name)
    return names

def safe_parse_json(text: str) -> dict:
    try:
        return json.loads(text)
    except:
        return {
            'professional_title': 'N/A',
            'institution': 'N/A',
            'city_state': 'N/A',
            'bio': 'N/A'
        }

def parse_sessions(llm_output: str) -> list[dict]:
    sessions = []
    pattern = r"\*\*Title:\*\*\s*(.*?)\n\*\*Date:\*\*\s*(.*?)\n\*\*Start Time\s*-\s*End Time:\*\*\s*(.*?)\n\*\*Location:\*\*\s*(.*?)\n\*\*Presenter\(s\):\*\*\s*(.*?)\n\*\*Description:\*\*\s*(.*?)(?=\n\n|\Z)"
    matches = re.findall(pattern, llm_output, re.DOTALL)
    for match in matches:
        title, date, time_range, location, presenters, description = match
        sessions.append({
            'title': title.strip(),
            'date': date.strip(),
            'time': time_range.strip(),
            'location': location.strip(),
            'presenters': presenters.strip(),
            'description': description.strip()
        })
    return sessions

def add_hyperlink(paragraph, url, text, color = '0000FF', underline = True):
    part = paragraph.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external = True)
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')

    if color:
        c = OxmlElement('w:color')
        c.set(qn('w:val'), color)
        rPr.append(c)
    if underline:
        u = OxmlElement('w:u')
        u.set(qn('w:val'), 'single')
        rPr.append(u)
    new_run.append(rPr)
    t = OxmlElement('w:t')
    t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink

def set_table_border(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), 'auto')
        tblBorders.append(border)
    tblPr.append(tblBorders)

# ─────────────────────────────────────────────────────────────────────────────
# LLM and tools setup
# ─────────────────────────────────────────────────────────────────────────────

llm = ChatGoogleGenerativeAI(
    model = 'gemini-2.0-flash',
    google_api_key = os.getenv('GOOGLE_API_KEY'),
    temperature = 0.3
)

tools = [tavily_search_tool, scrape_conference_website_tool, linkedin_search_tool, table_tool]
tools_dict = {t.name: t for t in tools}
llm = llm.bind_tools(tools)

# ─────────────────────────────────────────────────────────────────────────────
# Streamlit app
# ─────────────────────────────────────────────────────────────────────────────

def main():
    st.set_page_config(page_title = 'Conference Planner', layout = 'wide')
    st.title('Conference Planner')

    st.session_state.setdefault('step', 'ask_conference')
    st.session_state.setdefault('conference_query', None)
    st.session_state.setdefault('interests', None)
    st.session_state.setdefault('intermediate_llm_output', None)
    st.session_state.setdefault('llm_output', None)
    st.session_state.setdefault('sessions', [])
    st.session_state.setdefault('selected_sessions', [])

    step = st.session_state.step

    if step == 'ask_conference':
        with st.chat_message('assistant'):
            st.write('What pharmaceutical conference are you interested in?')
        conf = st.chat_input('Enter the name of a conference:')
        if conf:
            st.session_state.conference_query = conf
            st.session_state.step = 'ask_interest'
            st.rerun()
        return
    
    if step == 'ask_interest':
        with st.chat_message('assistant'):
            st.write('What topics or types of sessions are you interested in?')
        interest = st.chat_input('Enter your interests:')
        if interest:
            st.session_state.interests = interest
            st.session_state.step = 'llm_process'
            st.rerun()
        return

    if step == 'llm_process':
        with st.spinner('Searching for sessions and analyzing content...'):
            try:
                conf = st.session_state.conference_query
                interest = st.session_state.interests

                st.write(f'Searching for conference: {conf}')
                url = tools_dict['tavily_search_tool'].invoke({'query': conf})
                st.write(f'Tavily returned URL: {url}')

                if not url.startswith('http'):
                    raise Exception(f'Tavily failed: {url}')
                
                scraped_text = tools_dict['scrape_conference_website_tool'].invoke({'url': url})
                print(scraped_text[:1000])
                prompt = f'''
You are a skilled research assistant. Extract sessions related to "{interest}" from the {conf} pharmaceutical conference.

Here is the scraped website content (limit: 5000 characters):

{scraped_text}

Please extract up to 10 relevant sessions that match the user's interests.

Return the results as a clean Markdown list like this:

**Title:** Example  
**Date:** July 30  
**Start Time - End Time:** 10:00 AM - 11:00 AM  
**Location:** Room X  
**Presenter(s):** Jane Doe  
**Description:** Talk description here.

Separate each session with a blank line.

If fields are missing, explain why.

Return the results as a clean, readable bulleted list.
'''
                result = llm.invoke(prompt)
                st.session_state.intermediate_llm_output = result.content
                st.session_state.step = 'linkedin_search'
            except Exception as e:
                st.error(f'LLM processing failed: {e}')        
        st.rerun()
        return
    
    if step == 'linkedin_search':
        with st.spinner('Searching for LinkedIn profiles...'):
            try:
                raw_output = st.session_state.intermediate_llm_output
                sessions = parse_sessions(raw_output)
                enriched_sessions = []

                for session in sessions:
                    presenters = extract_presenter_names(f"**Presenter(s):** {session['presenters']}")
                    profile_data = []
                    for name in presenters:
                        enriched = tools_dict['linkedin_search_tool'].invoke({'name_and_affiliation': name})
                        profile_data.append(enriched)
                    if profile_data:
                        session['enriched_presenters'] = profile_data
                        session['professional_titles'] = ': '.join(p['professional_title'] for p in profile_data)
                        session['institution'] = ': '.join(p['institution'] for p in profile_data)
                        session['city_state'] = ': '.join(p['city_state'] for p in profile_data)
                        session['bio'] = '\n\n'.join(p['bio'] for p in profile_data)
                        session['linkedin_url'] = profile_data[0]['linkedin_url']
                    else:
                        session['professional_titles'] = 'N/A'
                        session['institution'] = 'N/A'
                        session['city_state'] = 'N/A'
                        session['bio'] = 'N/A'
                        session['linkedin_url'] = 'N/A'
                    enriched_sessions.append(session)
                st.session_state.sessions = enriched_sessions
                st.session_state.step = 'show_result'
            except Exception as e:
                st.error(f'LinkedIn lookup failed: {e}')
        st.rerun()
        return
    
    if step == 'show_result':
        with st.chat_message('assistant'):
            raw_output = st.session_state.intermediate_llm_output
            sessions = parse_sessions(raw_output)
            st.session_state.sessions = sessions

            st.write('Select sessions you would like to attend:')
            selected_sessions = []
            for i, session in enumerate(sessions):
                with st.expander(f"{session['title']}"):
                    st.markdown(f"**Date:** {session['date']}")
                    st.markdown(f"**Time:** {session['time']}")
                    st.markdown(f"**Location:** {session['location']}")
                    presenters = session.get('enriched_presenters')
                    if presenters:
                        presenter_lines = []
                        for p in presenters:
                            name = p.get('name', 'Unknown')
                            linkedin = p.get('linkedin_url', 'search_bar_url')
                            if linkedin != 'N/A':
                                presenter_lines.append(f'[{name}] ({linkedin})')
                            else:
                                presenter_lines.append(name)
                        st.markdown(f"**Presenters:** "+", ".join(presenter_lines))
                    else:
                        st.markdown(f"**Presenters:** {session['presenters']}")
                    st.markdown(f"**Description:** {session['description']}")
                    if st.checkbox('Add this session', key = f'select_{i}'):
                        selected_sessions.append(session)
            
            st.session_state.selected_sessions = selected_sessions
            if selected_sessions:
                if st.button('Generate Table from Selected Sessions'):
                    st.session_state.step = 'generate_table'
                    st.rerun()
        return

    if step == 'generate_table':
        selected = st.session_state.get('selected_sessions', [])
        if not selected:
            st.waring('No sessions selected.')
            return

        table_input = []
        for s in selected:
            table_input.append({
                'names': s['presenters'],
                'professional_titles': s.get('professional_titles', 'N/A'),
                'institution': s.get('institution', 'N/A'),
                'city_state': s.get('city_state', 'N/A'),
                'title': s['title'],
                'dtl': f"{s['date']}, {s['time']}, @ {s['location']}",
                'bio': s.get('bio', 'N/A'),
                'linkedin_url': s.get('linkedin_url', 'N/A')
            })
        
        result = tools_dict['table_tool'].invoke({'info': table_input})
        st.success(result)

        with open('table.docx', 'rb') as file:
            st.download_button('Download Table', data = file, file_name = 'conference_table.docx')

        st.session_state.step = 'done'
        return
    
    if step == 'done':
        with st.chat_message('assistant'):
            st.write('Thank you for using the Conference Planner!')
        return
    
if __name__ == '__main__':
    main()
