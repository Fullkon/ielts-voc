# -*- coding: utf-8 -*-
"""
IELTS Vocabulary Learning App - 28 Day Study Plan
Streamlit application for learning 3,484 IELTS vocabulary words.
"""

import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime, date, timedelta
import json
import random
import re
from collections import defaultdict
from typing import List, Dict, Optional, Tuple


# =============================================================================
# CONFIGURATION
# =============================================================================

st.set_page_config(
    page_title="IELTS Vocabulary Trainer for Victor",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

TOTAL_DAYS = 28
DATA_FILE = Path(__file__).parent / 'progress.json'
TEST_BANK_FILE = Path(__file__).parent / 'test_bank.json'
TEST_BANK_DIR  = Path(__file__).parent / 'test_bank_llm'
XLSX_FILE = Path(__file__).parent / '雅思英文词汇表（完整版）.xlsx'
AUDIO_DIR = Path(__file__).parent / 'audio'


# =============================================================================
# CUSTOM CSS
# =============================================================================

def apply_custom_css():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    * { font-family: 'Inter', sans-serif; }
    
    .main-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
        border-radius: 16px;
        color: white;
        margin-bottom: 1.5rem;
        text-align: center;
    }
    .main-header h1 { font-size: 2.2rem; font-weight: 700; margin: 0; }
    .main-header p { font-size: 1rem; opacity: 0.9; margin-top: 0.3rem; }
    
    .stat-card {
        background: white;
        border-radius: 12px;
        padding: 1.2rem;
        box-shadow: 0 2px 12px rgba(0,0,0,0.08);
        text-align: center;
        border-left: 4px solid #667eea;
    }
    .stat-card .value { font-size: 2rem; font-weight: 700; color: #667eea; }
    .stat-card .label { font-size: 0.85rem; color: #888; margin-top: 0.3rem; }
    
    .word-card {
        background: white;
        border-radius: 12px;
        padding: 1.5rem;
        margin: 0.8rem 0;
        box-shadow: 0 2px 10px rgba(0,0,0,0.06);
        border-left: 4px solid #667eea;
        transition: transform 0.2s;
    }
    .word-card:hover { transform: translateY(-2px); box-shadow: 0 4px 16px rgba(0,0,0,0.1); }
    .word-card .word-en { font-size: 1.4rem; font-weight: 700; color: #333; }
    .word-card .word-cn { font-size: 1.1rem; color: #667eea; font-weight: 600; }
    .word-card .label-badge { 
        display: inline-block; 
        background: #f0f0ff; 
        color: #667eea; 
        padding: 0.15rem 0.6rem; 
        border-radius: 20px; 
        font-size: 0.75rem; 
        font-weight: 600; 
    }
    .word-card .content { color: #555; font-size: 0.9rem; margin-top: 0.5rem; line-height: 1.6; }
    
    .test-card {
        background: white;
        border-radius: 16px;
        padding: 2rem;
        box-shadow: 0 4px 20px rgba(0,0,0,0.08);
        margin: 1rem 0;
        text-align: center;
    }
    .test-card .question { font-size: 1.8rem; font-weight: 700; color: #333; margin-bottom: 1rem; }
    .test-card .hint { color: #999; font-size: 0.9rem; margin-bottom: 1rem; }
    
    .correct-badge {
        background: #e8f5e9;
        color: #2e7d32;
        padding: 0.5rem 1rem;
        border-radius: 8px;
        font-weight: 600;
        display: inline-block;
    }
    .wrong-badge {
        background: #ffebee;
        color: #c62828;
        padding: 0.5rem 1rem;
        border-radius: 8px;
        font-weight: 600;
        display: inline-block;
    }
    
    .review-item {
        background: #fff3e0;
        border-left: 4px solid #ff9800;
        border-radius: 8px;
        padding: 0.8rem 1rem;
        margin: 0.4rem 0;
    }
    
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #f8f9ff 0%, #f0f0ff 100%);
    }
    
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px 8px 0 0;
        padding: 10px 20px;
        font-weight: 600;
    }
    </style>
    """, unsafe_allow_html=True)


# =============================================================================
# DATA LOADER
# =============================================================================

@st.cache_data
def load_vocabulary() -> pd.DataFrame:
    df = pd.read_excel(XLSX_FILE)
    df.columns = ['index_no', 'word', 'pronunciation', 'english_def',
                  'collocations', 'sentence', 'root_words', 'related_words', 'chinese_def']
    df = df.dropna(subset=['word', 'chinese_def'])
    df['word'] = df['word'].astype(str).str.strip()
    df['chinese_def'] = df['chinese_def'].astype(str).str.strip()
    df['english_def'] = df['english_def'].fillna('').astype(str)
    df['pronunciation'] = df['pronunciation'].fillna('').astype(str)
    df['collocations'] = df['collocations'].fillna('').astype(str)
    df['sentence'] = df['sentence'].fillna('').astype(str)
    df['root_words'] = df['root_words'].fillna('').astype(str)
    df['related_words'] = df['related_words'].fillna('').astype(str)
    df = df.reset_index(drop=True)
    df['id'] = df.index
    return df

def parse_collocations(colloc_str: str) -> List[str]:
    """Parse comma- or pipe-separated collocation PHRASES.
    Each item is a full phrase like 'conduct research', NOT a single word.
    """
    if not colloc_str or colloc_str in ('None', 'nan', ''):
        return []
    parts = re.split(r'[,|，]', colloc_str)
    return [p.strip() for p in parts if p.strip()]

def parse_related_words(related_str: str) -> List[str]:
    if not related_str or related_str in ('None', 'nan', ''):
        return []
    parts = re.split(r'[,|，]', related_str)
    return [p.strip() for p in parts if p.strip()]

def parse_root_words(root_str: str) -> List[str]:
    if not root_str or root_str in ('None', 'nan', ''):
        return []
    parts = re.split(r'[,|，]', root_str)
    return [p.strip() for p in parts if p.strip()]

def get_audio_html(word_index: int) -> str:
    """Return compact HTML audio element with controls.
    Uses base64 data-URI so it works without separate file hosting.
    Streamlit preserves <audio controls> tags (but strips onclick/js handlers).
    """
    import base64
    audio_path = AUDIO_DIR / f'{word_index}.mp3'
    if not audio_path.exists() or audio_path.stat().st_size == 0:
        return ''
    try:
        with open(audio_path, 'rb') as f:
            b64 = base64.b64encode(f.read()).decode()
        return (
            f'<audio controls preload="none" '
            f'style="height:20px;width:140px;display:inline-block;vertical-align:middle;margin-left:6px;" '
            f'title="Click play to listen">'
            f'<source src="data:audio/mp3;base64,{b64}" type="audio/mp3">'
            f'</audio>'
        )
    except Exception:
        return ''


def render_audio_player(word_index: int):
    """Render an audio player for the word using st.audio."""
    audio_path = AUDIO_DIR / f'{word_index}.mp3'
    if audio_path.exists() and audio_path.stat().st_size > 0:
        try:
            with open(audio_path, 'rb') as f:
                audio_bytes = f.read()
            st.audio(audio_bytes, format='audio/mp3')
        except Exception:
            pass


# =============================================================================
# LEARNING PLAN GENERATOR
# =============================================================================

def generate_learning_plan(total_words: int) -> Dict[int, List[int]]:
    """Generate learning plan in original Excel order (no shuffle)."""
    plan = {}
    indices = list(range(total_words))
    
    base_per_day = total_words // TOTAL_DAYS
    remainder = total_words % TOTAL_DAYS
    
    start = 0
    for day in range(1, TOTAL_DAYS + 1):
        count = base_per_day + (1 if day <= remainder else 0)
        plan[day] = indices[start:start + count]
        start += count
    
    return plan


# =============================================================================
# TEST BANK MANAGER - Pre-generate & persist test questions
# =============================================================================

class TestBankManager:
    """Loads pre-generated LLM test questions from test_bank_llm/ directory.
    Each day has a day_DD.json file with 4 question arrays: cn2en, en2cn, collocation, sentence.
    """
    TEST_TYPES = ['cn2en', 'en2cn', 'collocation', 'sentence']

    def __init__(self):
        self._cache: dict[str, dict] = {}  # day_key -> bank dict

    def _load_day(self, day: int) -> dict | None:
        day_key = str(day)
        if day_key in self._cache:
            return self._cache[day_key]
        f = TEST_BANK_DIR / f"day_{day:02d}.json"
        if not f.exists():
            return None
        try:
            data = json.loads(f.read_text(encoding='utf-8'))
            self._cache[day_key] = data
            return data
        except Exception:
            return None

    def get_test_batch(self, day: int, test_type: str, size: int | None = None) -> list[dict]:
        """Return a list of question dicts for the given day + test type."""
        bank = self._load_day(day)
        if not bank:
            return []
        questions = bank.get("questions", {}).get(test_type, [])
        if size and size < len(questions):
            return questions[:size]
        return questions

    def has_day(self, day: int) -> bool:
        return (TEST_BANK_DIR / f"day_{day:02d}.json").exists()


# =============================================================================
# PROGRESS MANAGER
# =============================================================================

class ProgressManager:
    def __init__(self):
        self.data = self._load()
    
    def _load(self) -> dict:
        if DATA_FILE.exists():
            try:
                with open(DATA_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                pass
        return {
            'start_date': None,
            'daily_progress': {},
            'word_records': {},
            'test_history': [],
            'review_queue': [],
        }
    
    def save(self):
        DATA_FILE.parent.mkdir(parents=True, exist_ok=True)
        with open(DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)
    
    def init_start_date(self):
        if self.data['start_date'] is None:
            self.data['start_date'] = date.today().isoformat()
            self.save()
    
    def get_start_date(self) -> Optional[date]:
        if self.data['start_date']:
            return date.fromisoformat(self.data['start_date'])
        return None
    
    def get_day_progress(self, day: int) -> dict:
        return self.data['daily_progress'].get(str(day), {
            'completed': False,
            'score': 0,
            'total_tested': 0,
            'correct': 0,
            'time_spent': 0
        })
    
    def update_day_progress(self, day: int, score: float, total_tested: int, correct: int):
        prev = self.get_day_progress(day)
        self.data['daily_progress'][str(day)] = {
            'completed': prev.get('completed', False) or (total_tested > 0),
            'score': round(score, 1),
            'total_tested': total_tested,
            'correct': correct,
            'time_spent': prev.get('time_spent', 0),
            'last_updated': datetime.now().isoformat()
        }
        self.save()
    
    def get_word_record(self, word_id: int) -> dict:
        return self.data['word_records'].get(str(word_id), {
            'attempts': 0,
            'correct': 0,
            'test_scores': {'cn2en': 0, 'en2cn': 0, 'collocation': 0, 'sentence': 0},
            'mastery': 0.0,
            'last_tested': None,
            'first_seen': None,
            'review_count': 0
        })
    
    def update_word_record(self, word_id: int, test_type: str, correct: bool):
        key = str(word_id)
        rec = self.get_word_record(word_id)
        rec['attempts'] += 1
        if correct:
            rec['correct'] += 1
        n = rec['attempts']
        new_acc = rec['correct'] / n
        rec['test_scores'][test_type] = round(new_acc * 100)
        rec['mastery'] = round(new_acc, 2)
        rec['last_tested'] = datetime.now().isoformat()
        if rec['first_seen'] is None:
            rec['first_seen'] = datetime.now().isoformat()
        
        self.data['word_records'][key] = rec
        
        if rec['mastery'] < 0.6 and word_id not in self.data['review_queue']:
            self.data['review_queue'].append(word_id)
        elif rec['mastery'] >= 0.9 and word_id in self.data['review_queue']:
            self.data['review_queue'].remove(word_id)
        
        self.save()
    
    def add_test_history(self, day: int, test_type: str, score: float, total: int):
        self.data['test_history'].append({
            'date': date.today().isoformat(),
            'day': day,
            'type': test_type,
            'score': round(score, 1),
            'total': total
        })
        self.save()
    
    def get_review_queue(self) -> List[int]:
        return self.data['review_queue']
    
    def get_overall_stats(self) -> dict:
        recs = self.data['word_records']
        total_tested = len(recs)
        total_correct = sum(r['correct'] for r in recs.values())
        total_attempts = sum(r['attempts'] for r in recs.values())
        mastered = sum(1 for r in recs.values() if r['mastery'] >= 0.9)
        need_review = sum(1 for r in recs.values() if 0.3 <= r['mastery'] < 0.6)
        weak = sum(1 for r in recs.values() if r['mastery'] < 0.3 and r['attempts'] > 0)
        completed_days = sum(1 for d in self.data['daily_progress'].values() if d.get('completed', False))
        
        return {
            'total_tested': total_tested,
            'total_attempts': total_attempts,
            'accuracy': round(total_correct / total_attempts * 100, 1) if total_attempts > 0 else 0,
            'mastered': mastered,
            'need_review': need_review,
            'weak': weak,
            'completed_days': completed_days,
            'review_queue_size': len(self.data['review_queue'])
        }
    
    def get_weak_words(self, limit: int = 50) -> List[int]:
        recs = self.data['word_records']
        tested = [(int(k), v['mastery'], v['attempts']) for k, v in recs.items() if v['attempts'] > 0]
        tested.sort(key=lambda x: (x[1], -x[2]))
        return [t[0] for t in tested[:limit]]


# =============================================================================
# TEST ENGINE
# =============================================================================

class TestEngine:
    
    @staticmethod
    def get_question(question_item: dict, word_row: pd.Series, test_type: str) -> dict:
        """Wrap a pre-generated question item with UI metadata."""
        type_labels = {
            'cn2en':       '汉译英',
            'en2cn':       '英译汉',
            'collocation': '短语搭配翻译',
            'sentence':    '句子词汇填空',
        }
        return {
            'type':      test_type,
            'label':     type_labels.get(test_type, test_type),
            'question':  question_item.get('question', ''),
            'answer':    question_item.get('answer', '').strip().lower(),
            'hint':      question_item.get('hint', ''),
            'word_id':   question_item.get('word_id', 0),
        }
    
    @staticmethod
    def evaluate(user_answer: str, correct_answer: str) -> Tuple[bool, int]:
        user = user_answer.strip().lower()
        correct = correct_answer.strip().lower()
        
        if user == correct:
            return True, 100
        
        if len(user) > 2 and len(correct) > 2:
            if abs(len(user) - len(correct)) <= 2:
                distance = TestEngine._levenshtein(user, correct)
                max_len = max(len(user), len(correct))
                similarity = (1 - distance / max_len) * 100
                if similarity >= 85:
                    return True, int(similarity)
            if correct in user or user in correct:
                return True, 70
        
        return False, 0
    
    @staticmethod
    def _levenshtein(s1: str, s2: str) -> int:
        if len(s1) < len(s2):
            return TestEngine._levenshtein(s2, s1)
        if len(s2) == 0:
            return len(s1)
        prev = list(range(len(s2) + 1))
        for i, c1 in enumerate(s1):
            curr = [i + 1]
            for j, c2 in enumerate(s2):
                cost = 0 if c1 == c2 else 1
                curr.append(min(curr[j] + 1, prev[j + 1] + 1, prev[j] + cost))
            prev = curr
        return prev[-1]


# =============================================================================
# MAIN APPLICATION
# =============================================================================

def main():
    apply_custom_css()
    
    # --- Load Data ---
    df = load_vocabulary()
    total_words = len(df)
    plan = generate_learning_plan(total_words)
    
    # --- Initialize Session State ---
    if 'progress' not in st.session_state:
        st.session_state.progress = ProgressManager()
        st.session_state.progress.init_start_date()
    
    if 'test_bank' not in st.session_state:
        st.session_state.test_bank = TestBankManager()
    
    if 'test_batch' not in st.session_state:
        st.session_state.test_batch = []
    if 'test_index' not in st.session_state:
        st.session_state.test_index = 0
    if 'test_score' not in st.session_state:
        st.session_state.test_score = 0
    if 'test_total' not in st.session_state:
        st.session_state.test_total = 0
    if 'test_type' not in st.session_state:
        st.session_state.test_type = 'cn2en'
    if 'test_results' not in st.session_state:
        st.session_state.test_results = []
    if 'show_answer' not in st.session_state:
        st.session_state.show_answer = False
    if 'last_user_answer' not in st.session_state:
        st.session_state.last_user_answer = ''
    if 'learning_day' not in st.session_state:
        st.session_state.learning_day = 1
    if 'learning_page' not in st.session_state:
        st.session_state.learning_page = 0
    if 'test_batch_size' not in st.session_state:
        st.session_state.test_batch_size = 20
    if 'show_test_bottom_notice' not in st.session_state:
        st.session_state.show_test_bottom_notice = False
    
    pm = st.session_state.progress
    tbm = st.session_state.test_bank
    
    # =========================================================================
    # SIDEBAR
    # =========================================================================
    with st.sidebar:
        st.markdown("## 📚 IELTS词汇训练")
        st.markdown("---")
        
        st.markdown("### 📅 选择学习日")
        selected_day = st.selectbox(
            "学习日 (1-28)",
            options=list(range(1, TOTAL_DAYS + 1)),
            index=st.session_state.learning_day - 1,
            key="day_selector"
        )
        if selected_day != st.session_state.learning_day:
            st.session_state.learning_day = selected_day
            st.session_state.learning_page = 0
            st.session_state.test_batch = []
            st.session_state.test_index = 0
            st.rerun()
        
        day_words = plan[selected_day]
        day_progress = pm.get_day_progress(selected_day)
        
        st.markdown(f"**第 {selected_day} 天**: {len(day_words)} 个单词")
        if day_progress['completed']:
            st.markdown(f"✅ 已完成 | 正确率: {day_progress['score']:.0f}%")
        else:
            st.markdown("⏳ 待学习")
        
        st.markdown("---")
        
        st.markdown("### 📊 整体进度")
        stats = pm.get_overall_stats()
        
        col1, col2 = st.columns(2)
        with col1:
            st.metric("已完成天数", f"{stats['completed_days']}/28")
            st.metric("已测试词汇", stats['total_tested'])
        with col2:
            st.metric("总体正确率", f"{stats['accuracy']}%")
            st.metric("已掌握", f"{stats['mastered']} 词")
        
        progress_pct = stats['completed_days'] / TOTAL_DAYS
        st.progress(progress_pct, text=f"学习进度: {progress_pct*100:.0f}%")
        
        st.markdown("---")
        st.markdown("### 🎯 词汇状态")
        st.markdown(f"- ⭐ 已掌握: **{stats['mastered']}** 词")
        st.markdown(f"- 🔄 需复习: **{stats['need_review']}** 词")
        st.markdown(f"- ⚠️ 薄弱词: **{stats['weak']}** 词")
        st.markdown(f"- 📝 复习队列: **{stats['review_queue_size']}** 词")
        
        st.markdown("---")
        st.caption(f"开始日期: {pm.get_start_date()}")
        if st.button("🔄 重置进度", type="secondary"):
            if st.session_state.get('confirm_reset'):
                pm.data = {
                    'start_date': None,
                    'daily_progress': {},
                    'word_records': {},
                    'test_history': [],
                    'review_queue': []
                }
                pm.save()
                st.session_state.clear()
                st.rerun()
            else:
                st.session_state.confirm_reset = True
                st.warning("再次点击确认重置所有学习进度")
    
    # =========================================================================
    # MAIN CONTENT
    # =========================================================================
    
    st.markdown("""
    <div class="main-header">
        <h1>📖 IELTS 词汇训练营--Victor特制</h1>
        <p>28天系统攻克 3,484 核心雅思词汇 | 多模式测试 | 智能复习</p>
    </div>
    """, unsafe_allow_html=True)
    
    tab1, tab2, tab3, tab4 = st.tabs([
        "📚 今日学习",
        "✍️ 开始测试",
        "📊 学习报告",
        "🎯 复习强化"
    ])
    
    # =========================================================================
    # TAB 1: TODAY'S LEARNING
    # =========================================================================
    with tab1:
        st.markdown(f"## 📚 第 {selected_day} 天 - 词汇学习 ({len(day_words)} 词)")
        
        # Pagination
        page_size = 20
        total_pages = max(1, (len(day_words) + page_size - 1) // page_size)
        
        col_nav1, col_nav2, col_nav3 = st.columns([1, 2, 1])
        with col_nav1:
            if st.button("◀ 上一页", disabled=st.session_state.learning_page == 0):
                st.session_state.learning_page -= 1
                st.rerun()
        with col_nav2:
            st.markdown(f"<div style='text-align:center;padding-top:5px;'>第 {st.session_state.learning_page+1}/{total_pages} 页</div>", unsafe_allow_html=True)
        with col_nav3:
            if st.button("下一页 ▶", disabled=st.session_state.learning_page >= total_pages - 1):
                st.session_state.learning_page += 1
                st.rerun()
        
        # Display words
        start_idx = st.session_state.learning_page * page_size
        end_idx = min(start_idx + page_size, len(day_words))
        
        for i in range(start_idx, end_idx):
            word_idx = day_words[i]
            row = df.iloc[word_idx]
            word_rec = pm.get_word_record(word_idx)
            
            with st.container():
                # Word card header: word + pronunciation + audio (always visible)
                eng_def_preview = str(row['english_def'])[:250]
                eng_def_preview += '...' if len(str(row['english_def'])) > 250 else ''

                # Pronunciation from the dedicated column (IPA)
                pron_raw = str(row['pronunciation']).strip()
                pron_html = ""
                if pron_raw and pron_raw not in ('None', 'nan', ''):
                    pron_html = (
                        f'<span style="font-size:0.9rem;color:#888;margin-left:8px;'
                        f'font-style:italic;">{pron_raw}</span>'
                    )

                mastery_html = ""
                if word_rec['attempts'] > 0:
                    mastery_color = '#4caf50' if word_rec['mastery'] >= 0.8 else '#ff9800' if word_rec['mastery'] >= 0.5 else '#f44336'
                    mastery_html = (
                        f"<div style='margin-top:8px;'>"
                        f"<span style='color:{mastery_color};font-size:0.85rem;'>"
                        f"掌握度: {word_rec['mastery']*100:.0f}% | 测试: {word_rec['attempts']}次</span>"
                        f"</div>"
                    )

                # Render word card + audio side by side using st.columns for reliable audio
                card_col, audio_col = st.columns([0.88, 0.12])
                with card_col:
                    st.markdown(f"""
                    <div class="word-card">
                        <div style="display:flex;justify-content:space-between;align-items:center;">
                            <div>
                                <span class="word-en">{row['word']}</span>{pron_html}
                            </div>
                        </div>
                        <div class="content">
                            <p><span class="label-badge">英文释义</span> {eng_def_preview}</p>
                        </div>
                        {mastery_html}
                    </div>
                    """, unsafe_allow_html=True)
                with audio_col:
                    st.write("")  # vertical alignment spacer
                    render_audio_player(word_idx)

                # Expander: Chinese def, collocations (large), sentence, root, related
                with st.expander(f"点击查看: {row['word']} 的完整释义", expanded=False):
                    st.markdown(f"#### {row['word']} — {row['chinese_def']}")

                    colls = parse_collocations(str(row['collocations']))
                    # Post-process: if a collocation is a single word (no space, or doesn't
                    # contain the target word), treat it as an adjective/noun modifier and
                    # prepend/append the target word to form a proper collocation phrase.
                    processed_colls = []
                    target_word = str(row['word']).strip().lower()
                    for coll in colls:
                        cl = coll.strip()
                        if not cl:
                            continue
                        cl_lower = cl.lower()
                        # Already has target word? Keep as-is.
                        if target_word in cl_lower:
                            processed_colls.append(cl)
                        # Single word only — prepend target word to form a phrase
                        elif ' ' not in cl and ',' not in cl and '|' not in cl:
                            processed_colls.append(f"{target_word} {cl}")
                        else:
                            processed_colls.append(cl)
                    if processed_colls:
                        st.markdown("##### 常见搭配")
                        cols_disp = st.columns(min(4, len(processed_colls)))
                        for j, coll in enumerate(processed_colls):
                            with cols_disp[j % len(cols_disp)]:
                                st.markdown(f"<span style='font-size:1.05rem;font-weight:500;'>`{coll}`</span>", unsafe_allow_html=True)

                    if str(row['sentence']) and str(row['sentence']) not in ('None', 'nan', ''):
                        st.markdown(f"**例句**: *{row['sentence']}*")

                    root = parse_root_words(str(row['root_words']))
                    related = parse_related_words(str(row['related_words']))
                    if root:
                        st.markdown(f"**同根词**: {', '.join(root[:8])}")
                    if related:
                        st.markdown(f"**相近词**: {', '.join(related[:8])}")
                
                st.markdown("---")
        
        # Bottom pagination navigation
        col_bnav1, col_bnav2, col_bnav3 = st.columns([1, 2, 1])
        with col_bnav1:
            if st.button("◀ 上一页", key='bottom_prev', disabled=st.session_state.learning_page == 0):
                st.session_state.learning_page -= 1
                st.rerun()
        with col_bnav2:
            st.markdown(f"<div style='text-align:center;padding-top:5px;'>第 {st.session_state.learning_page+1}/{total_pages} 页</div>", unsafe_allow_html=True)
        with col_bnav3:
            if st.button("下一页 ▶", key='bottom_next', disabled=st.session_state.learning_page >= total_pages - 1):
                st.session_state.learning_page += 1
                st.rerun()
        
        # Direct test entry
        st.markdown("---")
        st.markdown("### ✍️ 直接进入测试")
        st.caption("测试题已在后台预生成并持久化，每次调用同一批题目，确保学习一致性。")
        
        col_t1, col_t2 = st.columns(2)
        with col_t1:
            test_type_go = st.selectbox(
                "选择测试题型",
                options=['cn2en', 'en2cn', 'collocation', 'sentence'],
                format_func=lambda x: {
                    'cn2en':       '汉译英',
                    'en2cn':       '英译汉',
                    'collocation': '短语搭配翻译',
                    'sentence':    '句子词汇填空',
                }[x],
                key='test_type_go_selector'
            )
        with col_t2:
            test_size_go = st.selectbox(
                "测试数量",
                options=[10, 20, 30, 50, "全部(约125题)"],
                index=1,  # default 20
                key='test_size_go_selector'
            )
        
        if st.button("🚀 进入测试", type="primary", use_container_width=True):
            # Load from persistent test bank
            size = len(day_words) if "全部" in str(test_size_go) else min(int(test_size_go), len(day_words))
            batch = tbm.get_test_batch(selected_day, test_type_go, size)
            if not batch:
                st.error(f"第{selected_day}天的测试题库尚未生成，请先运行 generate_test_bank.py")
            else:
                st.session_state.test_batch = batch
                st.session_state.test_index = 0
                st.session_state.test_score = 0
                st.session_state.test_total = 0
                st.session_state.test_results = []
                st.session_state.test_type = test_type_go
                st.session_state.test_batch_size = size
                st.session_state.show_answer = False
                st.session_state.last_user_answer = ''
                st.success(f"已加载 {len(batch)} 道测试题，请切换到「✍️ 开始测试」标签页")
        
        # Show "开始测试" button at bottom if test batch is loaded
        if st.session_state.test_batch and len(st.session_state.test_batch) > 0:
            st.markdown("---")
            st.markdown(f"""
            <div style="background:#e8f0fe;border-left:4px solid #667eea;border-radius:8px;padding:1rem;text-align:center;margin-bottom:1rem;">
                <p style="font-size:1.2rem;font-weight:700;color:#667eea;margin:0;">
                    测试已就绪！共 {len(st.session_state.test_batch)} 题 — 题型: {st.session_state.test_type}
                </p>
            </div>
            """, unsafe_allow_html=True)
            col_g1, col_g2, col_g3 = st.columns([1, 2, 1])
            with col_g2:
                if st.button("🎯 开始测试", key="bottom_start_test", type="primary", use_container_width=True):
                    st.session_state.show_test_bottom_notice = True
                    st.rerun()
                if st.session_state.get('show_test_bottom_notice', False):
                    st.info("请点击页面顶部的「✍️ 开始测试」标签页进入测试")
    
    # =========================================================================
    # TAB 2: TEST
    # =========================================================================
    with tab2:
        st.markdown("## ✍️ 词汇测试")
        
        if not st.session_state.test_batch:
            # Test bank auto-load display
            st.info("👈 请先在「📚 今日学习」页面选择题型并点击「进入测试」，或使用下方快速测试")
            
            st.markdown("---")
            st.markdown("### ⚡ 快速测试")
            quick_day = st.selectbox("选择天", list(range(1, TOTAL_DAYS+1)), key='quick_day')
            quick_size = st.selectbox("题数", [10, 20, 30], key='quick_size')
            quick_type = st.selectbox(
                "题型",
                options=['cn2en', 'en2cn', 'collocation', 'sentence'],
                format_func=lambda x: {
                    'cn2en': '汉译英', 'en2cn': '英译汉',
                    'collocation': '短语搭配翻译', 'sentence': '句子词汇填空',
                }[x],
                key='quick_type'
            )
            
            if st.button("⚡ 快速开始", type="primary", use_container_width=True):
                actual_size = min(quick_size, len(plan[quick_day]))
                batch = tbm.get_test_batch(quick_day, quick_type, actual_size)
                if not batch:
                    st.error(f"第{quick_day}天的测试题库尚未生成，请先运行 generate_test_bank.py")
                else:
                    st.session_state.test_batch = batch
                    st.session_state.test_index = 0
                    st.session_state.test_score = 0
                    st.session_state.test_total = 0
                    st.session_state.test_results = []
                    st.session_state.test_type = quick_type
                    st.session_state.show_answer = False
                    st.rerun()
        
        else:
            if st.session_state.test_index >= len(st.session_state.test_batch):
                # Test complete
                final_score = (st.session_state.test_score / st.session_state.test_total * 100) if st.session_state.test_total > 0 else 0
                st.balloons()
                st.markdown(f"""
                <div class="test-card">
                    <h2>🎉 测试完成!</h2>
                    <p style="font-size:3rem;font-weight:700;color:#667eea;">{final_score:.0f}%</p>
                    <p>正确: {st.session_state.test_score} / {st.session_state.test_total}</p>
                </div>
                """, unsafe_allow_html=True)
                
                pm.update_day_progress(selected_day, final_score, st.session_state.test_total, st.session_state.test_score)
                pm.add_test_history(selected_day, st.session_state.test_type, final_score, st.session_state.test_total)
                
                if st.session_state.test_results:
                    st.markdown("### 📋 详细结果")
                    for res in st.session_state.test_results[-20:]:
                        badge = '<span class="correct-badge">✓</span>' if res['correct'] else '<span class="wrong-badge">✗</span>'
                        st.markdown(
                            f"{badge} **{res['word']}** ({res['chinese_def']}) - 你的答案: `{res['user_answer']}`",
                            unsafe_allow_html=True
                        )
                        if not res['correct']:
                            st.caption(f"正确答案: {res['correct_answer']}")
                
                if st.button("🔄 返回学习", type="primary"):
                    st.session_state.test_batch = []
                    st.rerun()
            
            else:
                current_question = st.session_state.test_batch[st.session_state.test_index]
                test_type = st.session_state.test_type
                word_id = current_question.get('word_id', 0)
                word_row = df.iloc[word_id] if word_id < len(df) else df.iloc[0]

                question = TestEngine.get_question(current_question, word_row, test_type)

                progress_pct = st.session_state.test_index / len(st.session_state.test_batch)
                st.progress(progress_pct, text=f"进度: {st.session_state.test_index+1}/{len(st.session_state.test_batch)}")

                col_sc1, col_sc2 = st.columns(2)
                with col_sc1:
                    st.metric("正确", st.session_state.test_score)
                with col_sc2:
                    current_acc = (st.session_state.test_score / st.session_state.test_total * 100) if st.session_state.test_total > 0 else 0
                    st.metric("正确率", f"{current_acc:.0f}%")

                type_labels = {
                    'cn2en':       '汉译英',
                    'en2cn':       '英译汉',
                    'collocation': '短语搭配翻译',
                    'sentence':    '句子词汇填空',
                }

                st.markdown(f"""
                <div class="test-card">
                    <div style="color:#888;font-size:0.85rem;">{type_labels.get(test_type, '')}</div>
                    <div class="question">{question['question']}</div>
                    <div class="hint">{question.get('hint', '')}</div>
                </div>
                """, unsafe_allow_html=True)
                
                user_answer = st.text_input(
                    "请输入你的答案:",
                    key=f"answer_input_{st.session_state.test_index}",
                    placeholder="在此输入...",
                    disabled=st.session_state.show_answer
                )
                
                col_btn1, col_btn2, col_btn3 = st.columns(3)
                
                with col_btn1:
                    if st.button("✅ 提交", type="primary", use_container_width=True, disabled=st.session_state.show_answer):
                        if user_answer:
                            correct, similarity = TestEngine.evaluate(user_answer, question['answer'])
                            st.session_state.show_answer = True
                            st.session_state.last_user_answer = user_answer
                            st.session_state.test_total += 1
                            
                            if correct:
                                st.session_state.test_score += 1
                            
                            pm.update_word_record(word_id, test_type, correct)
                            st.session_state.test_results.append({
                                'word': word_row['word'],
                                'chinese_def': word_row['chinese_def'],
                                'user_answer': user_answer,
                                'correct_answer': question['answer'],
                                'correct': correct,
                                'similarity': similarity
                            })
                            st.rerun()
                        else:
                            st.warning("请输入答案")
                
                with col_btn2:
                    if st.button("💡 显示答案", use_container_width=True, disabled=st.session_state.show_answer):
                        st.session_state.show_answer = True
                        st.session_state.last_user_answer = '(跳过)'
                        st.session_state.test_total += 1
                        pm.update_word_record(word_id, test_type, False)
                        st.session_state.test_results.append({
                            'word': word_row['word'],
                            'chinese_def': word_row['chinese_def'],
                            'user_answer': '(跳过)',
                            'correct_answer': question['answer'],
                            'correct': False,
                            'similarity': 0
                        })
                        st.rerun()
                
                with col_btn3:
                    if st.button("➡️ 下一题", type="secondary", use_container_width=True):
                        st.session_state.test_index += 1
                        st.session_state.show_answer = False
                        st.session_state.last_user_answer = ''
                        st.rerun()
                
                if st.session_state.show_answer:
                    st.markdown("---")
                    correct_answer = question['answer']
                    user_ans = st.session_state.last_user_answer
                    
                    is_correct, _ = TestEngine.evaluate(user_ans, correct_answer)
                    
                    if is_correct and user_ans != '(跳过)':
                        st.markdown(f'<span class="correct-badge">✅ 正确!</span>', unsafe_allow_html=True)
                        st.markdown(f"你的答案: `{user_ans}` | 正确答案: `{correct_answer}`")
                    else:
                        st.markdown(f'<span class="wrong-badge">❌ 错误</span>', unsafe_allow_html=True)
                        st.markdown(f"你的答案: `{user_ans}`")
                        st.markdown(f"正确答案: **`{correct_answer}`**")
                    
                    st.markdown("---")
                    st.markdown(f"**📖 {word_row['word']}** — *{word_row['chinese_def']}*")
                    st.caption(str(word_row['english_def'])[:300])
                    
                    # Audio for the word in test answer
                    audio_path = AUDIO_DIR / f'{word_id}.mp3'
                    if audio_path.exists() and audio_path.stat().st_size > 0:
                        try:
                            with open(audio_path, 'rb') as f:
                                st.audio(f.read(), format='audio/mp3')
                        except Exception:
                            pass
                    
                    colls = parse_collocations(str(word_row['collocations']))
                    if colls:
                        st.markdown("**搭配**: " + " | ".join(f"`{c}`" for c in colls[:5]))
                    if str(word_row['sentence']) and str(word_row['sentence']) != 'None':
                        st.markdown(f"**例句**: {word_row['sentence']}")
    
    # =========================================================================
    # TAB 3: LEARNING REPORT
    # =========================================================================
    with tab3:
        st.markdown("## 📊 学习报告")
        
        stats = pm.get_overall_stats()
        
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            st.markdown(f"""
            <div class="stat-card">
                <div class="value">{stats['completed_days']}<span style="font-size:1rem">/28</span></div>
                <div class="label">已完成天数</div>
            </div>
            """, unsafe_allow_html=True)
        with col2:
            st.markdown(f"""
            <div class="stat-card" style="border-left-color:#4caf50;">
                <div class="value" style="color:#4caf50;">{stats['accuracy']}%</div>
                <div class="label">总体正确率</div>
            </div>
            """, unsafe_allow_html=True)
        with col3:
            st.markdown(f"""
            <div class="stat-card" style="border-left-color:#2196f3;">
                <div class="value" style="color:#2196f3;">{stats['total_tested']}</div>
                <div class="label">已测词汇</div>
            </div>
            """, unsafe_allow_html=True)
        with col4:
            st.markdown(f"""
            <div class="stat-card" style="border-left-color:#ff9800;">
                <div class="value" style="color:#ff9800;">{stats['mastered']}</div>
                <div class="label">已掌握</div>
            </div>
            """, unsafe_allow_html=True)
        with col5:
            st.markdown(f"""
            <div class="stat-card" style="border-left-color:#f44336;">
                <div class="value" style="color:#f44336;">{stats['need_review']}</div>
                <div class="label">需复习</div>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("---")
        st.markdown("### 📈 每日进度")
        days_data = []
        for d in range(1, TOTAL_DAYS + 1):
            dp = pm.get_day_progress(d)
            days_data.append({
                'Day': f'第{d}天',
                '完成': 1 if dp['completed'] else 0,
                '正确数': dp.get('correct', 0),
                '测试数': dp.get('total_tested', 0),
                '正确率': dp.get('score', 0)
            })
        
        df_progress = pd.DataFrame(days_data)
        
        col_chart1, col_chart2 = st.columns(2)
        with col_chart1:
            st.bar_chart(df_progress.set_index('Day')[['正确率']], use_container_width=True)
        with col_chart2:
            st.markdown("#### 完成状态")
            completed_days_list = [d for d in range(1, TOTAL_DAYS+1) if pm.get_day_progress(d)['completed']]
            if completed_days_list:
                st.markdown(f"✅ 已完成: 第 {', '.join(map(str, completed_days_list[:14]))}天")
                if len(completed_days_list) > 14:
                    st.markdown(f"✅ 第 {', '.join(map(str, completed_days_list[14:28]))}天")
        
        st.markdown("### 📋 每日详情")
        for d in range(1, TOTAL_DAYS + 1):
            dp = pm.get_day_progress(d)
            status = "✅" if dp['completed'] else "⬜"
            day_word_count = len(plan.get(d, []))
            
            if dp['completed']:
                st.markdown(
                    f"{status} **第{d}天** ({day_word_count}词) | "
                    f"正确率: {dp['score']:.0f}% | "
                    f"正确: {dp['correct']}/{dp['total_tested']}"
                )
            else:
                st.markdown(f"{status} 第{d}天 ({day_word_count}词) | 待学习")
        
        st.markdown("---")
        st.markdown("### 📝 测试历史")
        test_hist = pm.data['test_history']
        if test_hist:
            df_tests = pd.DataFrame(test_hist[-50:])
            st.dataframe(df_tests, use_container_width=True, hide_index=True)
        else:
            st.info("暂无测试记录，快去测试吧!")
    
    # =========================================================================
    # TAB 4: REVIEW & REINFORCEMENT
    # =========================================================================
    with tab4:
        st.markdown("## 🎯 复习与强化训练")
        
        review_queue = pm.get_review_queue()
        weak_words = pm.get_weak_words(limit=100)
        
        col_r1, col_r2 = st.columns(2)
        with col_r1:
            st.metric("复习队列", len(review_queue))
        with col_r2:
            st.metric("薄弱词汇(Top100)", len(weak_words))
        
        st.markdown("---")
        
        review_mode = st.radio(
            "选择复习模式",
            options=['smart_review', 'weak_focus', 'random_review', 'daily_review', 'mastery_check'],
            format_func=lambda x: {
                'smart_review': '🧠 智能复习 (根据遗忘曲线推荐)',
                'weak_focus': '🎯 薄弱词强化',
                'random_review': '🎲 随机抽查',
                'daily_review': '📅 按天复习',
                'mastery_check': '✅ 掌握度检查'
            }[x],
            key='review_mode'
        )
        
        review_words = []
        review_label = ""
        
        if review_mode == 'smart_review':
            today_idx = selected_day
            review_indices = set(review_queue)
            for offset in [1, 3, 7, 14]:
                if today_idx - offset >= 1:
                    review_indices.update(plan.get(today_idx - offset, []))
            review_words = list(review_indices)[:50]
            review_label = "智能复习 (复习队列 + 间隔重复)"
        
        elif review_mode == 'weak_focus':
            review_words = weak_words[:50]
            review_label = "薄弱词强化 (掌握度从低到高)"
        
        elif review_mode == 'random_review':
            all_indices = list(range(total_words))
            review_words = random.sample(all_indices, min(50, total_words))
            review_label = "随机抽查 50 词"
        
        elif review_mode == 'daily_review':
            review_day = st.selectbox("选择要复习的天", list(range(1, TOTAL_DAYS+1)), key='review_day_select')
            review_words = plan.get(review_day, [])[:50]
            review_label = f"复习第{review_day}天词汇"
        
        elif review_mode == 'mastery_check':
            recs = pm.data['word_records']
            mid_mastery = [int(k) for k, v in recs.items() if 0.3 <= v['mastery'] < 0.85 and v['attempts'] > 0]
            random.shuffle(mid_mastery)
            review_words = mid_mastery[:50]
            review_label = "掌握度检查 (30%-85%掌握度)"
        
        st.markdown(f"**{review_label}**: {len(review_words)} 个单词")
        
        if review_words:
            st.markdown("---")
            
            batch_size = st.slider("每页显示数量", 5, 30, 10, key='review_batch')
            page_r = st.session_state.get('review_page', 0)
            total_review_pages = max(1, (len(review_words) + batch_size - 1) // batch_size)
            
            col_rnav1, col_rnav2, col_rnav3 = st.columns([1, 2, 1])
            with col_rnav1:
                if st.button("◀", key='review_prev', disabled=page_r == 0):
                    st.session_state.review_page = max(0, page_r - 1)
                    st.rerun()
            with col_rnav2:
                st.markdown(f"<div style='text-align:center;'>{page_r+1}/{total_review_pages}</div>", unsafe_allow_html=True)
            with col_rnav3:
                if st.button("▶", key='review_next', disabled=page_r >= total_review_pages - 1):
                    st.session_state.review_page = min(total_review_pages - 1, page_r + 1)
                    st.rerun()
            
            start_r = page_r * batch_size
            end_r = min(start_r + batch_size, len(review_words))
            
            for i in range(start_r, end_r):
                widx = review_words[i]
                row = df.iloc[widx]
                wrec = pm.get_word_record(widx)
                mastery = wrec['mastery']
                
                if mastery >= 0.8:
                    badge_color, status = '#4caf50', '🟢'
                elif mastery >= 0.5:
                    badge_color, status = '#ff9800', '🟡'
                elif mastery > 0:
                    badge_color, status = '#f44336', '🔴'
                else:
                    badge_color, status = '#999', '⚪'
                
                st.markdown(f"""
                <div class="review-item">
                    <div style="display:flex;justify-content:space-between;align-items:center;">
                        <div>
                            <strong style="font-size:1.1rem;">{row['word']}</strong>
                            <span style="color:#667eea;font-weight:600;"> — {row['chinese_def']}</span>
                        </div>
                        <div>
                            <span style="color:{badge_color};font-size:0.85rem;">{status} 掌握: {mastery*100:.0f}%</span>
                            <span style="color:#888;font-size:0.75rem;"> | 测试{wrec['attempts']}次</span>
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
                with st.expander(f"查看详情: {row['word']}"):
                    st.markdown(f"**英文释义**: {str(row['english_def'])[:300]}")
                    st.markdown(f"**汉语释义**: {row['chinese_def']}")
                    
                    # Audio for review word
                    audio_path = AUDIO_DIR / f'{widx}.mp3'
                    if audio_path.exists() and audio_path.stat().st_size > 0:
                        try:
                            with open(audio_path, 'rb') as f:
                                st.audio(f.read(), format='audio/mp3')
                        except Exception:
                            pass
                    
                    colls = parse_collocations(str(row['collocations']))
                    if colls:
                        st.markdown("**搭配**: " + " | ".join(f"`{c}`" for c in colls[:6]))
                    
                    related = parse_related_words(str(row['related_words']))
                    root = parse_root_words(str(row['root_words']))
                    if related:
                        st.markdown(f"**相近词**: {', '.join(related[:6])}")
                    if root:
                        st.markdown(f"**同根词**: {', '.join(root[:6])}")
                    
                    if st.button(f"⚡ 快速测试此词", key=f"qt_{widx}"):
                        row = df.iloc[widx]
                        word = str(row['word']).strip()
                        cn_def = str(row['chinese_def']).strip()
                        hint = f"{word[0].upper()}..." if len(word) > 3 else f"{len(word)} letters"
                        st.session_state.test_batch = [{
                            'word_id': widx,
                            'question': cn_def,
                            'answer': word.lower(),
                            'hint': hint,
                        }]
                        st.session_state.test_index = 0
                        st.session_state.test_score = 0
                        st.session_state.test_total = 0
                        st.session_state.test_results = []
                        st.session_state.show_answer = False
                        st.session_state.test_type = 'cn2en'
                        st.info("请切换到「✍️ 开始测试」标签页")
            
            st.markdown("---")
            review_test_count = min(30, len(review_words))
            if st.button(f"🚀 对这{review_test_count}个词进行测试", type="primary", use_container_width=True):
                test_words = review_words[:review_test_count]
                batch = []
                for widx in test_words:
                    row = df.iloc[widx]
                    word = str(row['word']).strip()
                    cn_def = str(row['chinese_def']).strip()
                    hint = f"{word[0].upper()}..." if len(word) > 3 else f"{len(word)} letters"
                    batch.append({
                        'word_id': widx,
                        'question': cn_def,
                        'answer': word.lower(),
                        'hint': hint,
                    })
                st.session_state.test_batch = batch
                st.session_state.test_index = 0
                st.session_state.test_score = 0
                st.session_state.test_total = 0
                st.session_state.test_results = []
                st.session_state.show_answer = False
                st.session_state.test_type = 'cn2en'
                st.info("请切换到「✍️ 开始测试」标签页")
        else:
            st.info("暂无需要复习的词汇。完成更多测试后，系统会自动识别薄弱词汇。")
        
        st.markdown("---")
        st.markdown("### 💡 学习建议")
        
        if stats['completed_days'] == 0:
            st.info("🔰 **开始建议**: 从第1天开始，先浏览词汇（点击下拉框查看汉语释义），再进入测试，每天坚持学习约125个单词。")
        elif stats['accuracy'] < 50:
            st.warning("⚠️ **准确率较低**: 建议多花时间浏览词汇的搭配和例句，不要急于看汉语释义，先尝试回忆再点击查看。")
        elif stats['accuracy'] < 70:
            st.info("📈 **稳步提升中**: 正确率在提高，继续保持！多利用搭配和句子加深理解。")
        elif stats['accuracy'] >= 70:
            st.success("🌟 **表现优秀**: 正确率很高！可以尝试搭配测试和句子填空等高级测试模式。")
        
        if stats['need_review'] > 50:
            st.warning(f"📋 有 {stats['need_review']} 个词汇需要复习，建议优先在「复习强化」页面进行薄弱词专项训练。")
        
        if stats['completed_days'] >= 3:
            st.markdown("### 🗓️ 间隔重复提醒 (基于艾宾浩斯遗忘曲线)")
            for off in [1, 3, 7, 14]:
                rev_day = selected_day - off
                if rev_day >= 1:
                    day_prog = pm.get_day_progress(rev_day)
                    if day_prog['completed']:
                        st.markdown(f"- **第{rev_day}天** ({off}天前): 🔄 建议复习 - 正确率 {day_prog['score']:.0f}%")


if __name__ == "__main__":
    main()
