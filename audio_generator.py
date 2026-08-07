# -*- coding: utf-8 -*-
"""
IELTS Vocabulary Audio Generator
Supports: edge-tts (free, primary) and Volcengine TTS (secondary)
"""

import asyncio
import os
import json
import time
import hashlib
import hmac
import uuid
from pathlib import Path
from typing import Optional
import requests

# =============================================================================
# CONFIGURATION
# =============================================================================

AUDIO_DIR = Path(__file__).parent / 'audio'
PROGRESS_FILE = AUDIO_DIR / 'generation_progress.json'

# edge-tts voice: clear American English female
EDGE_VOICE = 'en-US-JennyNeural'

# Volcengine TTS config (requires separate Voice Tech credentials)
VOLC_TTS_APPID = os.environ.get('VOLC_TTS_APPID', '')
VOLC_TTS_TOKEN = os.environ.get('VOLC_TTS_TOKEN', '')
VOLC_TTS_VOICE = 'en_female_garce_bigtts'  # English female voice

# Ark API key for potential Ark-based TTS
_ark_key_fallback = None
try:
    from config import ARK_KEY as _ark_key_fallback
except ImportError:
    pass
ARK_API_KEY = os.environ.get('ARK_API_KEY', _ark_key_fallback or '')

# Rate limiting
RATE_LIMIT_DELAY = 0.3  # seconds between requests


# =============================================================================
# AUDIO GENERATION ENGINE
# =============================================================================

class AudioGenerator:
    """Generate audio pronunciation files for vocabulary words."""

    def __init__(self):
        AUDIO_DIR.mkdir(parents=True, exist_ok=True)
        self.progress = self._load_progress()

    def _load_progress(self) -> dict:
        if PROGRESS_FILE.exists():
            try:
                with open(PROGRESS_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                pass
        return {'generated': {}, 'total': 0, 'last_idx': -1}

    def _save_progress(self):
        with open(PROGRESS_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.progress, f, ensure_ascii=False, indent=2)

    def get_audio_path(self, word_index: int) -> Path:
        return AUDIO_DIR / f'{word_index}.mp3'

    def has_audio(self, word_index: int) -> bool:
        path = self.get_audio_path(word_index)
        return path.exists() and path.stat().st_size > 0

    # -------------------------------------------------------------------------
    # edge-tts (primary - free, high quality)
    # -------------------------------------------------------------------------

    async def _generate_edge_tts(self, text: str, output_path: Path) -> bool:
        """Generate audio using Microsoft Edge TTS (free)."""
        try:
            import edge_tts
            communicate = edge_tts.Communicate(
                text=text,
                voice=EDGE_VOICE,
                rate='-5%',      # slightly slower for clarity
                pitch='+0Hz'
            )
            await communicate.save(str(output_path))
            return True
        except Exception as e:
            print(f'  edge-tts error: {e}')
            return False

    # -------------------------------------------------------------------------
    # Volcengine TTS (secondary - requires credentials)
    # -------------------------------------------------------------------------

    def _generate_volcengine_tts(self, text: str, output_path: Path) -> bool:
        """Generate audio using Volcengine TTS API."""
        if not VOLC_TTS_APPID or not VOLC_TTS_TOKEN:
            return False

        try:
            url = 'https://openspeech.bytedance.com/api/v1/tts'
            headers = {
                'Authorization': f'Bearer;{VOLC_TTS_TOKEN}',
                'Content-Type': 'application/json'
            }
            payload = {
                'app': {
                    'appid': VOLC_TTS_APPID,
                    'token': 'placeholder_token',
                    'cluster': 'volcano_tts'
                },
                'user': {'uid': 'ielts_vocab_app'},
                'audio': {
                    'voice_type': VOLC_TTS_VOICE,
                    'encoding': 'mp3',
                    'speed_ratio': 0.95,
                    'volume_ratio': 1.0,
                    'pitch_ratio': 1.0
                },
                'request': {
                    'reqid': str(uuid.uuid4()),
                    'text': text,
                    'text_type': 'plain',
                    'operation': 'query',
                    'with_frontend': 1
                }
            }

            resp = requests.post(url, json=payload, headers=headers, timeout=30)
            if resp.status_code == 200:
                with open(output_path, 'wb') as f:
                    f.write(resp.content)
                return True
            else:
                print(f'  Volcengine TTS error: {resp.status_code} - {resp.text[:200]}')
                return False
        except Exception as e:
            print(f'  Volcengine TTS error: {e}')
            return False

    # -------------------------------------------------------------------------
    # Ark-based TTS (newer API variant)
    # -------------------------------------------------------------------------

    def _generate_ark_tts(self, text: str, output_path: Path) -> bool:
        """Generate audio using Ark platform TTS endpoint."""
        try:
            url = 'https://ark.cn-beijing.volces.com/api/v3/audio/speech'
            headers = {
                'Authorization': f'Bearer {ARK_API_KEY}',
                'Content-Type': 'application/json'
            }
            payload = {
                'model': 'doubao-tts-1.0',  # TTS model
                'input': text,
                'voice': 'en_male_gentle',
                'response_format': 'mp3',
                'speed': 0.95
            }

            resp = requests.post(url, json=payload, headers=headers, timeout=30)
            if resp.status_code == 200:
                with open(output_path, 'wb') as f:
                    f.write(resp.content)
                return True
            elif resp.status_code in (401, 403, 404):
                print(f'  Ark TTS: endpoint not available or auth failed ({resp.status_code})')
                return False
            else:
                print(f'  Ark TTS error: {resp.status_code} - {resp.text[:200]}')
                return False
        except Exception as e:
            print(f'  Ark TTS error: {e}')
            return False

    # -------------------------------------------------------------------------
    # Main generation logic
    # -------------------------------------------------------------------------

    async def generate_one(self, word_index: int, word: str, force: bool = False) -> bool:
        """Generate audio for a single word. Returns True if successful."""
        output_path = self.get_audio_path(word_index)

        if not force and self.has_audio(word_index):
            return True

        text = word.strip()

        # Try engines in order: edge-tts -> Volcengine -> Ark
        engines = [
            ('edge-tts', lambda: self._generate_edge_tts(text, output_path)),
            ('volcengine', lambda: self._generate_volcengine_tts(text, output_path)),
            ('ark', lambda: self._generate_ark_tts(text, output_path)),
        ]

        for engine_name, engine_fn in engines:
            success = await engine_fn() if engine_name == 'edge-tts' else engine_fn()
            if success and self.has_audio(word_index):
                self.progress['generated'][str(word_index)] = {
                    'word': word,
                    'engine': engine_name,
                    'time': time.strftime('%Y-%m-%d %H:%M:%S')
                }
                self._save_progress()
                return True
            if self.has_audio(word_index):
                # edge-tts saves asynchronously, check again
                self.progress['generated'][str(word_index)] = {
                    'word': word,
                    'engine': engine_name,
                    'time': time.strftime('%Y-%m-%d %H:%M:%S')
                }
                self._save_progress()
                return True

        return False

    async def generate_batch(self, words: list, start_idx: int = 0,
                              end_idx: int = None, force: bool = False) -> dict:
        """Generate audio for a batch of words.
        
        Args:
            words: list of (index, word_text) tuples
            start_idx: starting position in the list
            end_idx: ending position (exclusive)
            force: regenerate even if exists
        
        Returns:
            dict with success/fail counts
        """
        if end_idx is None:
            end_idx = len(words)

        batch = words[start_idx:end_idx]
        total = len(batch)
        success = 0
        failed = 0

        print(f'Generating audio for {total} words (indices {start_idx}-{end_idx-1})...')

        for i, (word_idx, word_text) in enumerate(batch):
            if (i + 1) % 50 == 0:
                print(f'  Progress: {i+1}/{total} ({success} ok, {failed} fail)')

            try:
                ok = await self.generate_one(word_idx, word_text, force=force)
                if ok:
                    success += 1
                else:
                    failed += 1
            except Exception as e:
                print(f'  Error on word {word_idx} "{word_text}": {e}')
                failed += 1

            # Rate limiting
            await asyncio.sleep(RATE_LIMIT_DELAY)

        self.progress['total'] = len(self.progress['generated'])
        self.progress['last_idx'] = max(
            self.progress.get('last_idx', -1),
            end_idx - 1 if end_idx > 0 else -1
        )
        self._save_progress()

        print(f'Batch complete: {success} success, {failed} failed')
        return {'success': success, 'failed': failed, 'total': total}

    def get_stats(self) -> dict:
        """Get generation statistics."""
        generated = sum(1 for p in AUDIO_DIR.glob('*.mp3') if p.stat().st_size > 0)
        return {
            'generated': generated,
            'total_tracked': len(self.progress.get('generated', {})),
            'last_idx': self.progress.get('last_idx', -1)
        }


# =============================================================================
# COMMAND LINE INTERFACE
# =============================================================================

async def main():
    """Run audio generation from command line."""
    import pandas as pd

    xlsx = Path(__file__).parent / '雅思英文词汇表（完整版）.xlsx'
    if not xlsx.exists():
        print(f'Excel file not found: {xlsx}')
        return

    df = pd.read_excel(xlsx)
    df.columns = ['index_no', 'word', 'pronunciation', 'english_def',
                   'collocations', 'sentence', 'root_words', 'related_words', 'chinese_def']
    df['word'] = df['word'].astype(str).str.strip()

    generator = AudioGenerator()

    # Build word list
    words = [(i, str(df.iloc[i]['word']).strip()) for i in range(len(df))]

    stats = generator.get_stats()
    print(f'Already generated: {stats["generated"]} / {len(words)}')

    import sys
    force = '--force' in sys.argv
    batch_size = 100  # generate in batches to handle interruptions

    start = stats['last_idx'] + 1 if not force else 0
    if start >= len(words):
        print('All words already generated!')
        return

    print(f'Starting from index {start}...')

    for batch_start in range(start, len(words), batch_size):
        batch_end = min(batch_start + batch_size, len(words))
        result = await generator.generate_batch(
            words,
            start_idx=batch_start,
            end_idx=batch_end,
            force=force
        )
        print(f'  Batch {batch_start}-{batch_end-1}: {result}')

    final_stats = generator.get_stats()
    print(f'\nDone! Total generated: {final_stats["generated"]} / {len(words)}')


if __name__ == '__main__':
    asyncio.run(main())
