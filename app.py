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
    df.columns = ['index_no', 'word', 'english_def', 'chinese_def',
                   'collocations', 'sentence', 'root_words', 'related_words', 'notes']
    df = df.dropna(subset=['word', 'chinese_def'])
    df['word'] = df['word'].astype(str).str.strip()
    df['chinese_def'] = df['chinese_def'].astype(str).str.strip()
    df['english_def'] = df['english_def'].fillna('').astype(str)
    df['collocations'] = df['collocations'].fillna('').astype(str)
    df['sentence'] = df['sentence'].fillna('').astype(str)
    df['root_words'] = df['root_words'].fillna('').astype(str)
    df['related_words'] = df['related_words'].fillna('').astype(str)
    df['notes'] = df['notes'].fillna('').astype(str)
    df = df.reset_index(drop=True)
    df['id'] = df.index
    return df

def parse_collocations(colloc_str: str) -> List[str]:
    if not colloc_str or colloc_str == 'None':
        return []
    parts = re.split(r'[|,，;；]', colloc_str)
    return [p.strip() for p in parts if p.strip()]

def parse_related_words(related_str: str) -> List[str]:
    if not related_str or related_str == 'None':
        return []
    parts = re.split(r'[|,，;；]', related_str)
    return [p.strip() for p in parts if p.strip()]

def parse_root_words(root_str: str) -> List[str]:
    if not root_str or root_str == 'None':
        return []
    parts = re.split(r'[|,，;；]', root_str)
    return [p.strip() for p in parts if p.strip()]


# =============================================================================
# LEARNING PLAN GENERATOR
# =============================================================================

def generate_learning_plan(total_words: int) -> Dict[int, List[int]]:
    plan = {}
    indices = list(range(total_words))
    np.random.seed(42)
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
# TEST BANK MANAGER - Pre-generate & persist test questions
# =============================================================================

class TestBankManager:
    """Manages pre-generated test questions, persisted to disk."""
    
    def __init__(self, plan: Dict[int, List[int]]):
        self.plan = plan
        self.data = self._load()
    
    def _load(self) -> dict:
        if TEST_BANK_FILE.exists():
            try:
                with open(TEST_BANK_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                pass
        return {}
    
    def save(self):
        with open(TEST_BANK_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)
    
    def ensure_day_bank(self, day: int):
        """Ensure test bank exists for the given day, generate if not."""
        day_key = str(day)
        if day_key not in self.data:
            day_words = self.plan.get(day, [])
            self.data[day_key] = {}
            for test_type in ['translation', 'collocation', 'sentence', 'related']:
                shuffled = list(day_words)
                random.shuffle(shuffled)
                self.data[day_key][test_type] = shuffled
            self.save()
    
    def get_test_batch(self, day: int, test_type: str, size: int = None) -> List[int]:
        """Get pre-generated word IDs for the specified day and test type."""
        self.ensure_day_bank(day)
        day_key = str(day)
        word_ids = self.data[day_key].get(test_type, self.plan.get(day, []))
        if size and size < len(word_ids):
            return word_ids[:size]
        return word_ids
    
    def ensure_exists_for_day(self, day: int) -> bool:
        return str(day) in self.data


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
# RICH COLLOCATION GENERATOR (English only, multi-word)
# =============================================================================

VERB_NOUN = [
    'use', 'apply', 'develop', 'achieve', 'create', 'involve', 'consider', 'affect',
    'improve', 'understand', 'provide', 'require', 'support', 'include', 'produce',
    'establish', 'maintain', 'enhance', 'promote', 'identify', 'address', 'conduct',
    'implement', 'generate', 'assess', 'evaluate', 'facilitate', 'contribute to',
    'participate in', 'focus on', 'rely on', 'depend on', 'lead to', 'result in',
    'deal with', 'cope with', 'take into account', 'pay attention to',
    'make use of', 'have access to', 'play a role in', 'take part in',
    'make a contribution to', 'have an impact on', 'give rise to',
]

ADJ_NOUN = [
    'important', 'crucial', 'essential', 'relevant', 'significant', 'appropriate',
    'effective', 'potential', 'major', 'fundamental', 'key', 'primary', 'critical',
    'beneficial', 'comprehensive', 'efficient', 'valuable', 'necessary', 'common',
    'particular', 'sufficient', 'adequate', 'considerable', 'substantial',
    'widespread', 'prominent', 'remarkable', 'inevitable', 'inextricably linked',
]

NOUN_PREP = ['of', 'for', 'with', 'in', 'from', 'to', 'by', 'on', 'at', 'about', 'between', 'among']

MULTI_WORD_PATTERNS = [
    "the {adj} {noun}",
    "a {adj} {noun}",
    "{verb} the {noun}",
    "{noun} and {noun}",
    "in terms of {noun}",
    "in the context of {noun}",
    "a wide range of {noun}",
    "the impact of {noun} on",
    "as a result of {noun}",
    "with respect to {noun}",
    "in relation to {noun}",
    "a great deal of {noun}",
    "the vast majority of {noun}",
    "at the expense of {noun}",
    "take advantage of {noun}",
    "have a significant impact on {noun}",
    "play a crucial role in {noun}",
    "in the field of {noun}",
    "in the process of {noun}",
]


def generate_rich_collocations(word: str, existing: List[str]) -> dict:
    """Generate collocations grouped into 'existing' and 'extra', English only.
    Rules: max 1 preposition collocation, rest are content-word collocations, total extra ≤ 7."""
    existing_word = word.lower()
    seen = set()
    
    existing_colls = []
    for coll in existing:
        if coll and coll.strip() and coll.strip() not in seen:
            existing_colls.append(coll.strip())
            seen.add(coll.strip())
    
    extra_colls = []
    
    # 1. Verb + noun (content word)
    verb_count = min(4, max(1, 6 - len(existing)))
    selected_verbs = random.sample(VERB_NOUN[:30], min(verb_count, 30))
    for verb in selected_verbs:
        coll = f"{verb} {existing_word}"
        if coll not in seen and len(extra_colls) < 7:
            extra_colls.append(coll)
            seen.add(coll)
    
    # 2. Adjective + noun (content word)
    adj_count = min(3, max(1, 5 - len(existing)))
    selected_adj = random.sample(ADJ_NOUN[:25], min(adj_count, 25))
    for adj in selected_adj:
        coll = f"{adj} {existing_word}"
        if coll not in seen and len(extra_colls) < 7:
            extra_colls.append(coll)
            seen.add(coll)
    
    # 3. Noun + preposition — at most 1
    if len(extra_colls) < 7 and len(existing) < 4:
        prep = random.choice(NOUN_PREP)
        coll = f"{existing_word} {prep}"
        if coll not in seen:
            extra_colls.append(coll)
            seen.add(coll)
    
    # 4. Multi-word patterns (content word) — fill remaining slots up to 7
    if len(extra_colls) < 7:
        patterns = random.sample(MULTI_WORD_PATTERNS, min(7 - len(extra_colls), len(MULTI_WORD_PATTERNS)))
        adj_pool = ADJ_NOUN[:15]
        for pat in patterns:
            if len(extra_colls) >= 7:
                break
            adj = random.choice(adj_pool)
            noun = existing_word
            if "{adj}" in pat and "{noun}" in pat:
                coll = pat.replace("{adj}", adj).replace("{noun}", noun)
            elif "{adj}" in pat:
                coll = pat.replace("{adj}", adj)
            elif "{noun}" in pat:
                coll = pat.replace("{noun}", noun)
            elif "{verb}" in pat:
                coll = pat.replace("{verb}", random.choice(VERB_NOUN[:10]))
            else:
                continue
            if coll not in seen and coll != existing_word:
                extra_colls.append(coll)
                seen.add(coll)
    
    return {
        'existing': existing_colls,
        'extra': extra_colls
    }


# =============================================================================
# TEST ENGINE
# =============================================================================

class TestEngine:
    
    @staticmethod
    def generate_question(word_id: int, word_row: pd.Series, test_type: str, seed_offset: int = 0) -> dict:
        """Generate a deterministic question for the given word and test type."""
        rng = random.Random(word_id * 100 + seed_offset + {'translation':0,'collocation':1000,'sentence':2000,'related':3000}[test_type])
        
        word = str(word_row['word']).strip()
        word_lower = word.lower()
        
        if test_type == 'translation':
            hint = f"首字母: {word[0].upper()}..." if len(word) > 3 else f"长度: {len(word)} 个字母"
            return {
                'type': 'translation',
                'label': '汉译英',
                'question': str(word_row['chinese_def']),
                'answer': word_lower,
                'hint': hint,
                'word_id': int(word_id),
            }
        
        elif test_type == 'collocation':
            colls = parse_collocations(str(word_row['collocations']))
            if colls:
                # Use only the first collocation
                coll = colls[0].lower()
                # Strip parenthetical content (Chinese translations etc.)
                coll = re.sub(r'\s*[\(（][^)）]*[\)）]\s*', '', coll).strip()
                if word_lower in coll:
                    question = coll.replace(word_lower, '_____', 1)
                else:
                    question = f"_____ {coll}"
            else:
                if rng.random() > 0.5:
                    question = f"_____ {word_lower}"
                else:
                    question = f"{word_lower} _____"
            return {
                'type': 'collocation',
                'label': '搭配测试',
                'question': question,
                'answer': word_lower,
                'hint': f'目标词释义请自行回忆',
                'word_id': int(word_id),
            }
        
        elif test_type == 'sentence':
            sentence = str(word_row['sentence'])
            if sentence and sentence != 'None':
                pattern = re.compile(re.escape(word), re.IGNORECASE)
                question = pattern.sub('_____', sentence, count=1)
            else:
                question = f'翻译: {word_row["chinese_def"]}'
            return {
                'type': 'sentence',
                'label': '句子填空',
                'question': question,
                'answer': word_lower,
                'hint': f'长度: {len(word)} 个字母',
                'word_id': int(word_id),
            }
        
        elif test_type == 'related':
            related = parse_related_words(str(word_row['related_words']))
            root = parse_root_words(str(word_row['root_words']))
            clues = []
            if related:
                clues.append(f"同类词: {', '.join(rng.sample(related, min(4, len(related))))}")
            if root:
                clues.append(f"同根词: {', '.join(rng.sample(root, min(4, len(root))))}")
            if not clues:
                clues.append(f"释义: {str(word_row['chinese_def'])}")
            return {
                'type': 'related',
                'label': '相关词联想',
                'question': ' | '.join(clues),
                'answer': word_lower,
                'hint': '根据相关词/同根词回忆目标单词',
                'word_id': int(word_id),
            }
        
        # fallback
        return {
            'type': 'translation',
            'label': '汉译英',
            'question': str(word_row['chinese_def']),
            'answer': word_lower,
            'hint': '',
            'word_id': int(word_id),
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
        st.session_state.test_bank = TestBankManager(plan)
    
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
    if 'test_batch_size' not in st.session_state:
        st.session_state.test_batch_size = 20
    
    pm = st.session_state.progress
    tbm = st.session_state.test_bank
    
    # Ensure test bank exists for current day
    tbm.ensure_day_bank(st.session_state.learning_day)
    
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
            # Ensure bank for the new day
            tbm.ensure_day_bank(selected_day)
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
                # Also reset test bank
                tbm.data = {}
                tbm.save()
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
                # Word card header: English word + pronunciation + English definition visible
                eng_def_preview = str(row['english_def'])[:250]
                eng_def_preview += '...' if len(str(row['english_def'])) > 250 else ''
                
                # Pronunciation from notes column (IPA)
                notes_val = str(row['notes']).strip()
                pron_html = ""
                if notes_val and notes_val != 'None' and notes_val:
                    # Clean up IPA text - it may contain escape chars
                    pron_text = notes_val.replace('/', '').strip()
                    pron_html = f'<span style="font-size:0.85rem;color:#888;margin-left:8px;">/{pron_text}/</span>'
                
                mastery_html = ""
                if word_rec['attempts'] > 0:
                    mastery_color = '#4caf50' if word_rec['mastery'] >= 0.8 else '#ff9800' if word_rec['mastery'] >= 0.5 else '#f44336'
                    mastery_html = (
                        f"<div style='margin-top:8px;'>"
                        f"<span style='color:{mastery_color};font-size:0.85rem;'>"
                        f"掌握度: {word_rec['mastery']*100:.0f}% | 测试: {word_rec['attempts']}次</span>"
                        f"</div>"
                    )
                
                st.markdown(f"""
                <div class="word-card">
                    <div style="display:flex;justify-content:space-between;align-items:center;">
                        <div>
                            <span class="word-en">🔤 {row['word']}</span>{pron_html}
                        </div>
                    </div>
                    <div class="content">
                        <p><span class="label-badge">英文释义</span> {eng_def_preview}</p>
                    </div>
                    {mastery_html}
                </div>
                """, unsafe_allow_html=True)
                
                # Chinese definition + collocations + sentence + related/root words: ALL in expander
                with st.expander(f"💡 点击查看: {row['word']} 的汉语释义 & 搭配 & 例句", expanded=False):
                    st.markdown(f"**🇨🇳 汉语释义**: **{row['chinese_def']}**")
                    
                    # Existing collocations (English only)
                    existing_colls = parse_collocations(str(row['collocations']))
                    all_colls = generate_rich_collocations(row['word'], existing_colls)
                    
                    if all_colls['existing']:
                        st.markdown("**📎 常见搭配**:")
                        cols_disp = st.columns(min(4, len(all_colls['existing'])))
                        for j, coll in enumerate(all_colls['existing']):
                            with cols_disp[j % len(cols_disp)]:
                                st.markdown(f"`{coll}`")
                    
                    if all_colls['extra']:
                        st.markdown("**✨ 更多搭配**:")
                        for coll in all_colls['extra']:
                            st.markdown(f"- `{coll}`")
                    
                    # Sentence
                    if str(row['sentence']) and str(row['sentence']) != 'None':
                        st.markdown(f"**📝 例句**: *{row['sentence']}*")
                    
                    # Related & root words
                    related = parse_related_words(str(row['related_words']))
                    root = parse_root_words(str(row['root_words']))
                    if related:
                        st.markdown(f"**🔗 同类词**: {', '.join(related[:8])}")
                    if root:
                        st.markdown(f"**🌱 同根词**: {', '.join(root[:8])}")
                
                st.markdown("---")
        
        # Direct test entry - no generation step, bank already exists
        st.markdown("---")
        st.markdown("### ✍️ 直接进入测试")
        st.caption("测试题已在后台预生成并持久化，每次调用同一批题目，确保学习一致性。")
        
        col_t1, col_t2 = st.columns(2)
        with col_t1:
            test_type_go = st.selectbox(
                "选择测试题型",
                options=['translation', 'collocation', 'sentence', 'related'],
                format_func=lambda x: {
                    'translation': '🈂️ 汉译英',
                    'collocation': '🔗 搭配测试',
                    'sentence': '📝 句子填空',
                    'related': '🔍 相关词测试'
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
            st.session_state.test_batch = batch
            st.session_state.test_index = 0
            st.session_state.test_score = 0
            st.session_state.test_total = 0
            st.session_state.test_results = []
            st.session_state.test_type = test_type_go
            st.session_state.show_answer = False
            st.session_state.last_user_answer = ''
            st.session_state.test_batch_size = size
            st.success(f"已加载 {len(batch)} 道测试题，请切换到「✍️ 开始测试」标签页")
    
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
                tbm.ensure_day_bank(quick_day)
                batch = tbm.get_test_batch(quick_day, quick_type, actual_size)
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
                current_word_idx = st.session_state.test_batch[st.session_state.test_index]
                word_row = df.iloc[current_word_idx]
                test_type = st.session_state.test_type
                
                # Generate question deterministically from test bank position
                question = TestEngine.generate_question(
                    current_word_idx, word_row, test_type,
                    seed_offset=st.session_state.test_index
                )
                
                progress_pct = st.session_state.test_index / len(st.session_state.test_batch)
                st.progress(progress_pct, text=f"进度: {st.session_state.test_index+1}/{len(st.session_state.test_batch)}")
                
                col_sc1, col_sc2 = st.columns(2)
                with col_sc1:
                    st.metric("正确", st.session_state.test_score)
                with col_sc2:
                    current_acc = (st.session_state.test_score / st.session_state.test_total * 100) if st.session_state.test_total > 0 else 0
                    st.metric("正确率", f"{current_acc:.0f}%")
                
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
                    
                    colls = parse_collocations(str(row['collocations']))
                    if colls:
                        st.markdown("**搭配**: " + " | ".join(f"`{c}`" for c in colls[:6]))
                    
                    related = parse_related_words(str(row['related_words']))
                    root = parse_root_words(str(row['root_words']))
                    if related:
                        st.markdown(f"**同类词**: {', '.join(related[:6])}")
                    if root:
                        st.markdown(f"**同根词**: {', '.join(root[:6])}")
                    
                    if st.button(f"⚡ 快速测试此词", key=f"qt_{widx}"):
                        st.session_state.test_batch = [widx]
                        st.session_state.test_index = 0
                        st.session_state.test_score = 0
                        st.session_state.test_total = 0
                        st.session_state.test_results = []
                        st.session_state.show_answer = False
                        st.session_state.test_type = 'translation'
                        st.info("请切换到「✍️ 开始测试」标签页")
            
            st.markdown("---")
            review_test_count = min(30, len(review_words))
            if st.button(f"🚀 对这{review_test_count}个词进行测试", type="primary", use_container_width=True):
                test_words = review_words[:review_test_count]
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
