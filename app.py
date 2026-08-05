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
import hashlib
import copy


# =============================================================================
# CONFIGURATION
# =============================================================================

st.set_page_config(
    page_title="IELTS Vocabulary Trainer",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

TOTAL_DAYS = 28
DATA_FILE = Path(__file__).parent / 'progress.json'
XLSX_FILE = Path(__file__).parent / '雅思英文词汇表（完整版）.xlsx'


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
    
    .progress-bar-bg {
        width: 100%; height: 8px; background: #eee; border-radius: 4px; overflow: hidden;
    }
    .progress-bar-fg {
        height: 100%; background: linear-gradient(90deg, #667eea, #764ba2); border-radius: 4px;
    }
    
    .review-item {
        background: #fff3e0;
        border-left: 4px solid #ff9800;
        border-radius: 8px;
        padding: 0.8rem 1rem;
        margin: 0.4rem 0;
    }
    
    /* sidebar */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #f8f9ff 0%, #f0f0ff 100%);
    }
    
    /* tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
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
    """Load the Excel vocabulary file and return a clean DataFrame."""
    df = pd.read_excel(XLSX_FILE)
    # Rename columns to English for easier processing
    df.columns = ['index_no', 'word', 'english_def', 'chinese_def',
                   'collocations', 'sentence', 'root_words', 'related_words', 'notes']
    # Clean data
    df = df.dropna(subset=['word', 'chinese_def'])
    df['word'] = df['word'].astype(str).str.strip()
    df['chinese_def'] = df['chinese_def'].astype(str).str.strip()
    df['english_def'] = df['english_def'].fillna('').astype(str)
    df['collocations'] = df['collocations'].fillna('').astype(str)
    df['sentence'] = df['sentence'].fillna('').astype(str)
    df['root_words'] = df['root_words'].fillna('').astype(str)
    df['related_words'] = df['related_words'].fillna('').astype(str)
    df['notes'] = df['notes'].fillna('').astype(str)
    # Reset index
    df = df.reset_index(drop=True)
    df['id'] = df.index
    return df

def parse_collocations(colloc_str: str) -> List[str]:
    """Parse collocation string like 'crucial magma | use magma | magma of' into list."""
    if not colloc_str or colloc_str == 'None':
        return []
    parts = re.split(r'[|,，;；]', colloc_str)
    return [p.strip() for p in parts if p.strip()]

def parse_related_words(related_str: str) -> List[str]:
    """Parse related words string into list."""
    if not related_str or related_str == 'None':
        return []
    parts = re.split(r'[|,，;；]', related_str)
    return [p.strip() for p in parts if p.strip()]

def parse_root_words(root_str: str) -> List[str]:
    """Parse root words string into list."""
    if not root_str or root_str == 'None':
        return []
    parts = re.split(r'[|,，;；]', root_str)
    return [p.strip() for p in parts if p.strip()]


# =============================================================================
# LEARNING PLAN GENERATOR
# =============================================================================

def generate_learning_plan(total_words: int) -> Dict[int, List[int]]:
    """Generate 28-day learning plan. Returns {day: [word_indices]}."""
    plan = {}
    indices = list(range(total_words))
    np.random.seed(42)  # Fixed seed for reproducibility
    np.random.shuffle(indices)
    
    base_per_day = total_words // TOTAL_DAYS
    remainder = total_words % TOTAL_DAYS
    
    start = 0
    for day in range(1, TOTAL_DAYS + 1):
        count = base_per_day + (1 if day <= remainder else 0)
        plan[day] = indices[start:start + count]
        start += count
    
    return plan


# =============================================================================
# PROGRESS MANAGER
# =============================================================================

class ProgressManager:
    """Manages learning progress with JSON persistence."""
    
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
            'daily_progress': {},       # day -> {completed, score, time_spent}
            'word_records': {},         # word_id -> {attempts, correct, test_scores, mastery, last_tested}
            'test_history': [],         # [{date, day, type, score, total}]
            'review_queue': [],         # [word_id, ...]
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
        self.data['daily_progress'][str(day)] = {
            'completed': True,
            'score': round(score, 1),
            'total_tested': total_tested,
            'correct': correct,
            'time_spent': self.get_day_progress(day).get('time_spent', 0),
            'last_updated': datetime.now().isoformat()
        }
        self.save()
    
    def get_word_record(self, word_id: int) -> dict:
        return self.data['word_records'].get(str(word_id), {
            'attempts': 0,
            'correct': 0,
            'test_scores': {'translation': 0, 'collocation': 0, 'sentence': 0, 'related': 0},
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
        # Update test-specific score (binary for now)
        current = rec['test_scores'].get(test_type, 0)
        # Running average
        n = rec['attempts']
        old_acc = rec['correct'] / max(n-1, 1) if n > 1 else 0
        new_acc = rec['correct'] / n
        rec['test_scores'][test_type] = round(new_acc * 100)
        rec['mastery'] = round(new_acc, 2)
        rec['last_tested'] = datetime.now().isoformat()
        if rec['first_seen'] is None:
            rec['first_seen'] = datetime.now().isoformat()
        
        self.data['word_records'][key] = rec
        
        # Add to review queue if mastery < 0.6
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
        """Get IDs of weakest words for targeted practice."""
        recs = self.data['word_records']
        tested = [(int(k), v['mastery'], v['attempts']) for k, v in recs.items() if v['attempts'] > 0]
        tested.sort(key=lambda x: (x[1], -x[2]))  # Sort by mastery ascending, attempts descending
        return [t[0] for t in tested[:limit]]


# =============================================================================
# RICH COLLOCATION GENERATOR
# =============================================================================

COLLOCATION_TEMPLATES = {
    'verb': ['use', 'apply', 'develop', 'achieve', 'create', 'involve', 'consider', 'affect', 
             'improve', 'understand', 'provide', 'require', 'support', 'include', 'produce',
             'establish', 'maintain', 'enhance', 'promote', 'identify'],
    'adj': ['important', 'crucial', 'essential', 'relevant', 'significant', 'appropriate',
            'effective', 'potential', 'major', 'fundamental', 'key', 'primary', 'critical',
            'beneficial', 'comprehensive', 'efficient', 'valuable', 'necessary', 'common',
            'particular'],
    'prep': ['of', 'for', 'with', 'in', 'from', 'to', 'by', 'on', 'at', 'about', 'between', 'among'],
}

def generate_rich_collocations(word: str, existing: List[str]) -> List[Dict[str, str]]:
    """Generate rich collocations with Chinese translations."""
    rich = []
    existing_word = word.lower()
    
    # 1. Include existing collocations with translations
    for coll in existing:
        if coll:
            coll_clean = coll.strip()
            rich.append({
                'collocation': coll_clean,
                'type': 'origin',
                'cn': f"(原有搭配)"
            })
    
    # 2. Verb + Noun pattern
    verb_count = min(5, max(2, len(existing) // 2))
    selected_verbs = random.sample(COLLOCATION_TEMPLATES['verb'], verb_count)
    for verb in selected_verbs:
        coll = f"{verb} {existing_word}"
        if coll not in [r['collocation'] for r in rich]:
            rich.append({
                'collocation': coll,
                'type': 'verb_noun',
                'cn': f"动词+名词"
            })
    
    # 3. Adjective + Noun pattern
    adj_count = min(4, max(2, len(existing) // 3))
    selected_adj = random.sample(COLLOCATION_TEMPLATES['adj'], adj_count)
    for adj in selected_adj:
        coll = f"{adj} {existing_word}"
        if coll not in [r['collocation'] for r in rich]:
            rich.append({
                'collocation': coll,
                'type': 'adj_noun',
                'cn': f"形容词+名词"
            })
    
    # 4. Noun + Preposition pattern
    prep_count = min(3, max(1, len(existing) // 4))
    selected_prep = random.sample(COLLOCATION_TEMPLATES['prep'], prep_count)
    for prep in selected_prep:
        coll = f"{existing_word} {prep}"
        if coll not in [r['collocation'] for r in rich]:
            rich.append({
                'collocation': coll,
                'type': 'noun_prep',
                'cn': f"名词+介词"
            })
    
    return rich


# =============================================================================
# TEST ENGINE
# =============================================================================

class TestEngine:
    """Generates and evaluates various test types."""
    
    @staticmethod
    def generate_translation_test(word_row: pd.Series) -> dict:
        """Chinese → English translation test."""
        return {
            'type': 'translation',
            'label': '汉译英',
            'question': str(word_row['chinese_def']),
            'answer': str(word_row['word']).strip().lower(),
            'hint': f"首字母: {word_row['word'][0].upper()}..." if len(str(word_row['word'])) > 3 else f"长度: {len(str(word_row['word']))} 个字母",
            'word_row': word_row
        }
    
    @staticmethod
    def generate_collocation_test(word_row: pd.Series) -> dict:
        """Collocation fill-in-blank test."""
        coll_str = str(word_row['collocations'])
        colls = parse_collocations(coll_str)
        word = str(word_row['word']).strip()
        
        if not colls:
            # Fallback: create a simple collocation fill
            question = f"_____ {word}" if random.random() > 0.5 else f"{word} _____"
            return {
                'type': 'collocation',
                'label': '搭配测试',
                'question': question,
                'answer': '',
                'hint': '输入与目标词搭配的词',
                'word_row': word_row,
                'target_word': word,
                'show_word': True,
                'acceptable': []
            }
        
        # Pick a random collocation and hide the target word
        coll = random.choice(colls)
        coll_lower = coll.lower()
        word_lower = word.lower()
        
        if word_lower in coll_lower:
            blank_coll = coll_lower.replace(word_lower, '_____', 1)
        else:
            blank_coll = f"_____ {coll_lower}"
        
        return {
            'type': 'collocation',
            'label': '搭配测试',
            'question': blank_coll,
            'answer': word_lower,
            'hint': f'目标词汉语: {word_row["chinese_def"]}',
            'word_row': word_row,
            'target_word': word_lower,
            'show_word': False,
            'acceptable': []
        }
    
    @staticmethod
    def generate_sentence_test(word_row: pd.Series) -> dict:
        """Sentence fill-in-blank test."""
        sentence = str(word_row['sentence'])
        word = str(word_row['word']).strip()
        
        if not sentence or sentence == 'None':
            return {
                'type': 'sentence',
                'label': '句子测试',
                'question': f'翻译此句: {word_row["chinese_def"]}',
                'answer': word.lower(),
                'hint': f'目标词: {word}',
                'word_row': word_row,
                'target_word': word.lower()
            }
        
        # Create blank by replacing the word
        import re as re_mod
        pattern = re_mod.compile(re_mod.escape(word), re_mod.IGNORECASE)
        blank_sentence = pattern.sub('_____', sentence, count=1)
        
        return {
            'type': 'sentence',
            'label': '句子测试',
            'question': blank_sentence,
            'answer': word.lower(),
            'hint': f'汉语: {word_row["chinese_def"]} | 长度: {len(word)} 个字母',
            'word_row': word_row,
            'target_word': word.lower()
        }
    
    @staticmethod
    def generate_related_test(word_row: pd.Series) -> dict:
        """Related word recall test."""
        related = parse_related_words(str(word_row['related_words']))
        root = parse_root_words(str(word_row['root_words']))
        word = str(word_row['word']).strip()
        
        clues = []
        if related:
            clues.append(f"同类词: {', '.join(related[:4])}")
        if root:
            clues.append(f"同根词: {', '.join(root[:4])}")
        
        if not clues:
            clues.append(f"释义: {str(word_row['chinese_def'])}")
        
        return {
            'type': 'related',
            'label': '相关词测试',
            'question': ' | '.join(clues),
            'answer': word.lower(),
            'hint': f'根据相关词/同根词联想目标单词',
            'word_row': word_row,
            'target_word': word.lower(),
            'related_words': related,
            'root_words': root
        }
    
    @staticmethod
    def evaluate(user_answer: str, correct_answer: str) -> Tuple[bool, int]:
        """
        Evaluate user's answer against correct answer.
        Returns (is_correct, similarity_score 0-100).
        """
        user = user_answer.strip().lower()
        correct = correct_answer.strip().lower()
        
        if user == correct:
            return True, 100
        
        # Fuzzy matching for typos
        if len(user) > 2 and len(correct) > 2:
            # Levenshtein distance check
            if abs(len(user) - len(correct)) <= 2:
                distance = TestEngine._levenshtein(user, correct)
                max_len = max(len(user), len(correct))
                similarity = (1 - distance / max_len) * 100
                if similarity >= 85:
                    return True, int(similarity)
            
            # Partial match: correct word contains user's answer or vice versa
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
    
    if 'current_test' not in st.session_state:
        st.session_state.current_test = None
    if 'test_batch' not in st.session_state:
        st.session_state.test_batch = []
    if 'test_index' not in st.session_state:
        st.session_state.test_index = 0
    if 'test_score' not in st.session_state:
        st.session_state.test_score = 0
    if 'test_total' not in st.session_state:
        st.session_state.test_total = 0
    if 'test_type' not in st.session_state:
        st.session_state.test_type = 'translation'
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
    
    pm = st.session_state.progress
    
    # =========================================================================
    # SIDEBAR
    # =========================================================================
    with st.sidebar:
        st.markdown("## 📚 IELTS词汇训练")
        st.markdown("---")
        
        # Day selector
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
        
        # Day info
        day_words = plan[selected_day]
        day_progress = pm.get_day_progress(selected_day)
        
        st.markdown(f"**第 {selected_day} 天**: {len(day_words)} 个单词")
        if day_progress['completed']:
            st.markdown(f"✅ 已完成 | 正确率: {day_progress['score']:.0f}%")
        else:
            st.markdown("⏳ 待学习")
        
        st.markdown("---")
        
        # Overall stats
        st.markdown("### 📊 整体进度")
        stats = pm.get_overall_stats()
        
        col1, col2 = st.columns(2)
        with col1:
            st.metric("已完成天数", f"{stats['completed_days']}/28")
            st.metric("已测试词汇", stats['total_tested'])
        with col2:
            st.metric("总体正确率", f"{stats['accuracy']}%")
            st.metric("已掌握", f"{stats['mastered']} 词")
        
        # Progress bar
        progress_pct = stats['completed_days'] / TOTAL_DAYS
        st.progress(progress_pct, text=f"学习进度: {progress_pct*100:.0f}%")
        
        st.markdown("---")
        
        # Stats breakdown
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
    
    # Header
    st.markdown("""
    <div class="main-header">
        <h1>📖 IELTS 词汇训练营</h1>
        <p>28天系统攻克 3,484 核心雅思词汇 | 多模式测试 | 智能复习</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Tabs
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
        page_size = 10
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
                st.markdown(f"""
                <div class="word-card">
                    <div style="display:flex;justify-content:space-between;align-items:center;">
                        <span class="word-en">🔤 {row['word']}</span>
                        <span class="word-cn">{row['chinese_def']}</span>
                    </div>
                    <div class="content">
                        <p><span class="label-badge">释义</span> {str(row['english_def'])[:200]}{'...' if len(str(row['english_def'])) > 200 else ''}</p>
                </div>
                """, unsafe_allow_html=True)
                
                # Collocations section with rich content
                existing_colls = parse_collocations(str(row['collocations']))
                rich_colls = generate_rich_collocations(row['word'], existing_colls)
                
                # Display existing collocations first
                origin_colls = [c for c in rich_colls if c['type'] == 'origin']
                extra_colls = [c for c in rich_colls if c['type'] != 'origin']
                
                if origin_colls:
                    st.markdown("**📎 常见搭配**:")
                    cols = st.columns(min(4, len(origin_colls)))
                    for j, coll in enumerate(origin_colls):
                        with cols[j % len(cols)]:
                            st.markdown(f"`{coll['collocation']}` {coll['cn']}")
                
                # Rich / extra collocations (expandable)
                if extra_colls:
                    with st.expander(f"✨ 更多搭配 ({len(extra_colls)} 个) - 点击展开"):
                        for coll in extra_colls:
                            st.markdown(f"- `{coll['collocation']}` *({coll['cn']})*")
                
                # Sentence
                if str(row['sentence']) and str(row['sentence']) != 'None':
                    st.markdown(f"**📝 例句**: *{row['sentence']}*")
                
                # Related & root words
                related = parse_related_words(str(row['related_words']))
                root = parse_root_words(str(row['root_words']))
                if related or root:
                    rel_root_text = ""
                    if related:
                        rel_root_text += f"**🔗 同类词**: {', '.join(related[:5])}  "
                    if root:
                        rel_root_text += f"**🌱 同根词**: {', '.join(root[:5])}"
                    st.markdown(rel_root_text)
                
                # Mastery indicator
                if word_rec['attempts'] > 0:
                    mastery_color = '#4caf50' if word_rec['mastery'] >= 0.8 else '#ff9800' if word_rec['mastery'] >= 0.5 else '#f44336'
                    st.markdown(
                        f"<div style='margin-top:8px;'><span style='color:{mastery_color};font-size:0.85rem;'>"
                        f"掌握度: {word_rec['mastery']*100:.0f}% | 测试: {word_rec['attempts']}次</span></div>",
                        unsafe_allow_html=True
                    )
                
                st.markdown("</div>", unsafe_allow_html=True)
                st.markdown("---")
        
        # Start test section
        st.markdown("---")
        st.markdown("### ✍️ 开始测试")
        
        col_t1, col_t2 = st.columns(2)
        with col_t1:
            test_type_select = st.selectbox(
                "测试题型",
                options=['translation', 'collocation', 'sentence', 'related'],
                format_func=lambda x: {
                    'translation': '🈂️ 汉译英',
                    'collocation': '🔗 搭配测试',
                    'sentence': '📝 句子填空',
                    'related': '🔍 相关词测试'
                }[x],
                key='test_type_pre_select'
            )
        with col_t2:
            test_size = st.selectbox("测试数量", options=[10, 20, 30, 50, "全部(约125题)"], index=0, key='test_size_select')
        
        col_btn_a, col_btn_b = st.columns(2)
        with col_btn_a:
            if st.button(f"🚀 生成测试并开始", type="primary", use_container_width=True):
                size = len(day_words) if "全部" in str(test_size) else min(int(test_size), len(day_words))
                test_word_indices = random.sample(day_words, size)
                st.session_state.test_batch = test_word_indices
                st.session_state.test_index = 0
                st.session_state.test_score = 0
                st.session_state.test_total = 0
                st.session_state.test_results = []
                st.session_state.test_type = test_type_select
                st.session_state.show_answer = False
                st.session_state.last_user_answer = ''
                st.success(f"已生成 {size} 道测试题，请切换到「✍️ 开始测试」标签页进行测试")
        with col_btn_b:
            if st.button(f"📝 先学习全部词汇再测试", use_container_width=True):
                st.session_state.learning_page = 0
                st.info("已跳转到第1页，请浏览所有词汇后回来测试")
    
    # =========================================================================
    # TAB 2: TEST
    # =========================================================================
    with tab2:
        st.markdown("## ✍️ 词汇测试")
        
        # Test type selector
        if not st.session_state.test_batch:
            test_type = st.selectbox(
                "选择测试类型",
                options=['translation', 'collocation', 'sentence', 'related'],
                format_func=lambda x: {
                    'translation': '🈂️ 汉译英',
                    'collocation': '🔗 搭配测试',
                    'sentence': '📝 句子填空',
                    'related': '🔍 相关词测试'
                }[x],
                key="test_type_selector_empty"
            )
            st.session_state.test_type = test_type
            
            st.info("👈 请先在「📚 今日学习」页面生成测试题，或选择下方的复习模式")
            
            # Quick test - allow testing any day's words without going through learning
            st.markdown("---")
            st.markdown("### ⚡ 快速测试")
            quick_day = st.selectbox("选择天", list(range(1, TOTAL_DAYS+1)), key='quick_day')
            quick_size = st.selectbox("题数", [10, 20, 30], key='quick_size')
            quick_type = st.selectbox(
                "题型",
                options=['translation', 'collocation', 'sentence', 'related'],
                format_func=lambda x: {
                    'translation': '汉译英', 'collocation': '搭配测试',
                    'sentence': '句子填空', 'related': '相关词测试'
                }[x],
                key='quick_type'
            )
            
            if st.button("⚡ 快速开始", type="primary", use_container_width=True):
                qday_words = plan[quick_day]
                actual_size = min(quick_size, len(qday_words))
                st.session_state.test_batch = random.sample(qday_words, actual_size)
                st.session_state.test_index = 0
                st.session_state.test_score = 0
                st.session_state.test_total = 0
                st.session_state.test_results = []
                st.session_state.test_type = quick_type
                st.session_state.show_answer = False
                st.rerun()
        
        else:
            # Active test
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
                
                # Save progress
                pm.update_day_progress(selected_day, final_score, st.session_state.test_total, st.session_state.test_score)
                pm.add_test_history(selected_day, st.session_state.test_type, final_score, st.session_state.test_total)
                
                # Test results detail
                if st.session_state.test_results:
                    st.markdown("### 📋 详细结果")
                    for res in st.session_state.test_results[-20:]:  # Show last 20
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
                # Current test question
                current_word_idx = st.session_state.test_batch[st.session_state.test_index]
                word_row = df.iloc[current_word_idx]
                test_type = st.session_state.test_type
                
                # Generate question
                if test_type == 'translation':
                    question = TestEngine.generate_translation_test(word_row)
                elif test_type == 'collocation':
                    question = TestEngine.generate_collocation_test(word_row)
                elif test_type == 'sentence':
                    question = TestEngine.generate_sentence_test(word_row)
                elif test_type == 'related':
                    question = TestEngine.generate_related_test(word_row)
                else:
                    question = TestEngine.generate_translation_test(word_row)
                
                # Progress indicator
                progress_pct = st.session_state.test_index / len(st.session_state.test_batch)
                st.progress(progress_pct, text=f"进度: {st.session_state.test_index+1}/{len(st.session_state.test_batch)}")
                
                # Score display
                col_sc1, col_sc2 = st.columns(2)
                with col_sc1:
                    st.metric("正确", st.session_state.test_score)
                with col_sc2:
                    current_acc = (st.session_state.test_score / st.session_state.test_total * 100) if st.session_state.test_total > 0 else 0
                    st.metric("正确率", f"{current_acc:.0f}%")
                
                # Question card
                type_labels = {
                    'translation': '🈂️ 汉译英',
                    'collocation': '🔗 搭配测试',
                    'sentence': '📝 句子填空',
                    'related': '🔍 相关词测试'
                }
                
                st.markdown(f"""
                <div class="test-card">
                    <div style="color:#888;font-size:0.85rem;">{type_labels.get(test_type, '')}</div>
                    <div class="question">{question['question']}</div>
                    <div class="hint">💡 {question.get('hint', '')}</div>
                </div>
                """, unsafe_allow_html=True)
                
                # Answer input
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
                            
                            pm.update_word_record(current_word_idx, test_type, correct)
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
                        pm.update_word_record(current_word_idx, test_type, False)
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
                
                # Show answer result
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
                    
                    # Show word details
                    st.markdown("---")
                    st.markdown(f"**📖 {word_row['word']}** — *{word_row['chinese_def']}*")
                    st.caption(str(word_row['english_def'])[:300])
                    
                    # Show collocations
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
        
        # Top stats row
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
        
        # Daily progress chart
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
        
        # Day-by-day progress visualization
        col_chart1, col_chart2 = st.columns(2)
        with col_chart1:
            st.bar_chart(df_progress.set_index('Day')[['正确率']], use_container_width=True)
        with col_chart2:
            chart_data = df_progress.copy()
            chart_data['已完成'] = chart_data['完成'].map({1: 1, 0: 0})
            st.markdown("#### 完成状态")
            completed_days_list = [d for d in range(1, TOTAL_DAYS+1) if pm.get_day_progress(d)['completed']]
            if completed_days_list:
                st.markdown(f"✅ 已完成: 第 {', '.join(map(str, completed_days_list[:14]))}天")
                if len(completed_days_list) > 14:
                    st.markdown(f"✅ 第 {', '.join(map(str, completed_days_list[14:28]))}天")
        
        # Detailed day list
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
                st.markdown(
                    f"{status} 第{d}天 ({day_word_count}词) | 待学习",
                    help=f"共{day_word_count}个单词待学习"
                )
        
        # Test history
        st.markdown("---")
        st.markdown("### 📝 测试历史")
        test_hist = pm.data['test_history']
        if test_hist:
            df_tests = pd.DataFrame(test_hist[-50:])  # Last 50 tests
            st.dataframe(df_tests, use_container_width=True, hide_index=True)
        else:
            st.info("暂无测试记录，快去测试吧!")
    
    # =========================================================================
    # TAB 4: REVIEW & REINFORCEMENT
    # =========================================================================
    with tab4:
        st.markdown("## 🎯 复习与强化训练")
        
        # Review queue overview
        review_queue = pm.get_review_queue()
        weak_words = pm.get_weak_words(limit=100)
        
        col_r1, col_r2 = st.columns(2)
        with col_r1:
            st.metric("复习队列", len(review_queue))
        with col_r2:
            st.metric("薄弱词汇(Top100)", len(weak_words))
        
        st.markdown("---")
        
        # Review mode selector
        review_mode = st.radio(
            "选择复习模式",
            options=[
                'smart_review',
                'weak_focus',
                'random_review',
                'daily_review',
                'mastery_check'
            ],
            format_func=lambda x: {
                'smart_review': '🧠 智能复习 (根据遗忘曲线推荐)',
                'weak_focus': '🎯 薄弱词强化',
                'random_review': '🎲 随机抽查',
                'daily_review': '📅 按天复习',
                'mastery_check': '✅ 掌握度检查'
            }[x],
            key='review_mode'
        )
        
        # Determine review words
        review_words = []
        review_label = ""
        
        if review_mode == 'smart_review':
            # Smart review: mix of review queue + spaced repetition
            today_idx = selected_day
            review_indices = set(review_queue)
            # Add words from day (today-1), (today-3), (today-7), (today-14)
            for offset, label in [(1, "昨天"), (3, "3天前"), (7, "1周前"), (14, "2周前")]:
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
            # Words with mastery between 0.5 and 0.85
            recs = pm.data['word_records']
            mid_mastery = [int(k) for k, v in recs.items() if 0.3 <= v['mastery'] < 0.85 and v['attempts'] > 0]
            random.shuffle(mid_mastery)
            review_words = mid_mastery[:50]
            review_label = "掌握度检查 (30%-85%掌握度)"
        
        st.markdown(f"**{review_label}**: {len(review_words)} 个单词")
        
        if review_words:
            # Display review words
            st.markdown("---")
            
            # Batch display
            batch_size = st.slider("每页显示数量", 5, 30, 10, key='review_batch')
            page_r = st.session_state.get('review_page', 0)
            total_review_pages = max(1, (len(review_words) + batch_size - 1) // batch_size)
            
            col_rnav1, col_rnav2, col_rnav3 = st.columns([1, 2, 1])
            with col_rnav1:
                if st.button("◀", key='review_prev', disabled=page_r==0):
                    st.session_state.review_page = max(0, page_r - 1)
                    st.rerun()
            with col_rnav2:
                st.markdown(f"<div style='text-align:center;'>{page_r+1}/{total_review_pages}</div>", unsafe_allow_html=True)
            with col_rnav3:
                if st.button("▶", key='review_next', disabled=page_r>=total_review_pages-1):
                    st.session_state.review_page = min(total_review_pages-1, page_r + 1)
                    st.rerun()
            
            start_r = page_r * batch_size
            end_r = min(start_r + batch_size, len(review_words))
            
            for i in range(start_r, end_r):
                widx = review_words[i]
                row = df.iloc[widx]
                wrec = pm.get_word_record(widx)
                mastery = wrec['mastery']
                
                # Determine color based on mastery
                if mastery >= 0.8:
                    badge_color = '#4caf50'
                    status = '🟢'
                elif mastery >= 0.5:
                    badge_color = '#ff9800'
                    status = '🟡'
                elif mastery > 0:
                    badge_color = '#f44336'
                    status = '🔴'
                else:
                    badge_color = '#999'
                    status = '⚪'
                
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
                    st.markdown(f"**释义**: {str(row['english_def'])[:300]}")
                    
                    colls = parse_collocations(str(row['collocations']))
                    if colls:
                        st.markdown("**搭配**: " + " | ".join(f"`{c}`" for c in colls[:6]))
                    
                    related = parse_related_words(str(row['related_words']))
                    root = parse_root_words(str(row['root_words']))
                    if related:
                        st.markdown(f"**同类词**: {', '.join(related[:6])}")
                    if root:
                        st.markdown(f"**同根词**: {', '.join(root[:6])}")
                    
                    # Quick test button
                    if st.button(f"⚡ 快速测试此词", key=f"qt_{widx}"):
                        st.session_state.test_batch = [widx]
                        st.session_state.test_index = 0
                        st.session_state.test_score = 0
                        st.session_state.test_total = 0
                        st.session_state.test_results = []
                        st.session_state.show_answer = False
                        st.session_state.test_type = 'translation'
                        st.info("请切换到「✍️ 开始测试」标签页")
            
            # Start review test
            st.markdown("---")
            if st.button(f"🚀 对这{len(review_words[:30])}个词进行测试", type="primary", use_container_width=True):
                test_words = review_words[:30]
                st.session_state.test_batch = test_words
                st.session_state.test_index = 0
                st.session_state.test_score = 0
                st.session_state.test_total = 0
                st.session_state.test_results = []
                st.session_state.show_answer = False
                st.session_state.test_type = 'translation'
                st.info("请切换到「✍️ 开始测试」标签页")
        else:
            st.info("暂无需要复习的词汇。完成更多测试后，系统会自动识别薄弱词汇。")
        
        # Learning suggestions
        st.markdown("---")
        st.markdown("### 💡 学习建议")
        
        if stats['completed_days'] == 0:
            st.info("🔰 **开始建议**: 从第1天开始，先浏览词汇再进行测试，每天坚持学习约125个单词。")
        elif stats['accuracy'] < 50:
            st.warning("⚠️ **准确率较低**: 建议多花时间在「今日学习」页面上浏览词汇，熟悉后再测试。可以降低每天学习的词汇量，重在理解而非速度。")
        elif stats['accuracy'] < 70:
            st.info("📈 **稳步提升中**: 正确率在提高，继续保持！建议增加搭配学习和句子理解，有助于提升词汇的实际运用能力。")
        elif stats['accuracy'] >= 70:
            st.success("🌟 **表现优秀**: 正确率很高！可以加快学习进度，尝试更多搭配测试和句子填空等高级测试模式。")
        
        if stats['need_review'] > 50:
            st.warning(f"📋 有 {stats['need_review']} 个词汇需要复习，建议优先在「复习强化」页面进行薄弱词专项训练。")
        
        # Spaced repetition schedule reminder
        if stats['completed_days'] >= 3:
            st.markdown("### 🗓️ 间隔重复提醒 (基于艾宾浩斯遗忘曲线)")
            # Days that should be reviewed
            for off in [1, 3, 7, 14]:
                rev_day = selected_day - off
                if rev_day >= 1:
                    day_prog = pm.get_day_progress(rev_day)
                    if day_prog['completed']:
                        st.markdown(f"- **第{rev_day}天** ({off}天前): {'✅ 已复习' if False else '🔄 建议复习'} - 正确率 {day_prog['score']:.0f}%")


if __name__ == "__main__":
    main()
