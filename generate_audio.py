# -*- coding: utf-8 -*-
"""Batch audio generation script - optimized with concurrency."""
import asyncio
import sys
from pathlib import Path

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent))

from audio_generator import AudioGenerator, AUDIO_DIR
import pandas as pd

async def main():
    xlsx = Path(__file__).parent / '雅思英文词汇表（完整版）.xlsx'
    df = pd.read_excel(xlsx)
    df.columns = ['index_no', 'word', 'pronunciation', 'english_def',
                   'collocations', 'sentence', 'root_words', 'related_words', 'chinese_def']
    df['word'] = df['word'].astype(str).str.strip()
    
    generator = AudioGenerator()
    
    # Build word list
    words = [(i, str(df.iloc[i]['word']).strip()) for i in range(len(df))]
    total = len(words)
    
    # Check what's already generated
    stats = generator.get_stats()
    already = stats['generated']
    print(f'Total words: {total}')
    print(f'Already generated: {already}')
    print(f'Remaining: {total - already}')
    print()
    
    force = '--force' in sys.argv
    
    if not force and already >= total:
        print('All words already generated!')
        return
    
    # Generate in concurrent batches
    CONCURRENCY = 5  # parallel requests
    BATCH_SIZE = 200
    
    start_idx = 0 if force else stats.get('last_idx', -1) + 1
    if start_idx >= total:
        print('All words covered!')
        return
    
    print(f'Starting from index {start_idx}...')
    print(f'Concurrency: {CONCURRENCY}, Batch size: {BATCH_SIZE}')
    print()
    
    total_ok = 0
    total_fail = 0
    
    for batch_start in range(start_idx, total, BATCH_SIZE):
        batch_end = min(batch_start + BATCH_SIZE, total)
        batch = words[batch_start:batch_end]
        
        # Process in concurrent groups
        for i in range(0, len(batch), CONCURRENCY):
            chunk = batch[i:i + CONCURRENCY]
            tasks = []
            for word_idx, word_text in chunk:
                tasks.append(generator.generate_one(word_idx, word_text, force=force))
            
            results = await asyncio.gather(*tasks, return_exceptions=True)
            
            for j, result in enumerate(results):
                if isinstance(result, Exception):
                    total_fail += 1
                elif result:
                    total_ok += 1
                else:
                    total_fail += 1
            
            await asyncio.sleep(0.5)  # rate limiting between chunks
        
        progress = batch_end / total * 100
        print(f'Batch {batch_start}-{batch_end-1} done ({progress:.1f}%) | OK: {total_ok}, Fail: {total_fail}')
        
        # Save progress after each batch
        generator.progress['last_idx'] = batch_end - 1
        generator.progress['total'] = total_ok + already
        generator._save_progress()
    
    final_stats = generator.get_stats()
    print(f'\n{"="*50}')
    print(f'Generation complete!')
    print(f'Total generated: {final_stats["generated"]} / {total}')
    print(f'Success rate: {final_stats["generated"]/total*100:.1f}%')

if __name__ == '__main__':
    asyncio.run(main())
