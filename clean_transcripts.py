import re
import os
from pathlib import Path

# Set to True to restore the old, more destructive behaviour: any two long words
# where one contains the other count as duplicates, and every doubled word is
# collapsed to a single occurrence. This mangles legitimate text
# ("had had", "New York, New York", "state statement"), so it is off by default.
AGGRESSIVE = False

_INFLECTIONAL_SUFFIXES = ("ings", "ing", "ies", "es", "ed", "s")


def _stem(word):
    """Crude stemmer: strip a single inflectional suffix if enough stem remains."""
    for suffix in _INFLECTIONAL_SUFFIXES:
        if word.endswith(suffix) and len(word) - len(suffix) >= 4:
            return word[:-len(suffix)]
    return word


def _same_word(a, b):
    """True when two words are the same word, allowing for inflection.

    Unlike a plain substring test, this does not treat 'state'/'statement' or
    'bio'/'biology' as duplicates - only forms that share a stem, such as
    'process'/'processing' or 'study'/'studies'.
    """
    if a == b:
        return True
    if AGGRESSIVE:
        return len(a) > 4 and len(b) > 4 and (a in b or b in a)

    stem_a, stem_b = _stem(a), _stem(b)
    if len(stem_a) < 4 or len(stem_b) < 4:
        return False
    if stem_a == stem_b:
        return True
    # Tolerate a dropped 'e' ('use'/'using') or a doubled consonant ('run'/'running')
    return stem_a.rstrip('e') == stem_b.rstrip('e') or stem_a == stem_b[:-1] or stem_b == stem_a[:-1]


def remove_similar_word_sequences(text):
    """Remove runs of the same word repeated in different forms ('study studies studied')."""
    words = text.split()
    cleaned = []
    i = 0

    while i < len(words):
        current_word = words[i].lower().strip('.,;:!?')

        # Check for related words (same word, possibly inflected)
        j = i + 1
        similar_count = 0

        while j < len(words):
            next_word = words[j].lower().strip('.,;:!?')

            if _same_word(current_word, next_word):
                similar_count += 1
                j += 1
            else:
                break

        # If we found 2+ similar words, keep only the first one
        if similar_count >= 2:
            cleaned.append(words[i])
            print(f"  - Removed {similar_count} similar word(s) after '{words[i]}'")
            i = j
        else:
            cleaned.append(words[i])
            i += 1
    
    return ' '.join(cleaned)


def remove_repetitive_phrases(text, min_phrase_words=3):
    """
    Remove repetitive phrases from text using an improved algorithm.
    
    Args:
        text: Input text to clean
        min_phrase_words: Minimum number of words in a phrase to consider
    
    Returns:
        Cleaned text with repetitions removed
    """
    words = text.split()
    cleaned_words = []
    i = 0
    
    while i < len(words):
        # Try different phrase lengths, starting with longer phrases (up to 50 words)
        found_repetition = False
        
        for phrase_len in range(min(50, len(words) - i), min_phrase_words - 1, -1):
            if i + phrase_len * 2 > len(words):
                continue
                
            current_phrase = words[i:i + phrase_len]
            
            # Count consecutive repetitions
            repetition_count = 1
            j = i + phrase_len
            
            while j + phrase_len <= len(words):
                next_phrase = words[j:j + phrase_len]
                
                # Check if phrases match exactly
                if current_phrase == next_phrase:
                    repetition_count += 1
                    j += phrase_len
                else:
                    break
            
            # If we found ANY repetitions (even just 1 repeat), remove them
            if repetition_count >= 2:
                cleaned_words.extend(current_phrase)
                i = j
                found_repetition = True
                phrase_preview = ' '.join(current_phrase[:10])
                if len(current_phrase) > 10:
                    phrase_preview += '...'
                print(f"  - Removed {repetition_count - 1} repetition(s) of phrase ({phrase_len} words): '{phrase_preview}'")
                break
        
        if not found_repetition:
            cleaned_words.append(words[i])
            i += 1
    
    return ' '.join(cleaned_words)


def remove_word_repetitions(text, max_consecutive=2):
    """Remove excessive consecutive word repetitions."""
    words = text.split()
    cleaned = []
    
    i = 0
    while i < len(words):
        current_word = words[i]
        count = 1
        
        # Count consecutive occurrences
        while i + count < len(words) and words[i + count] == current_word:
            count += 1
        
        # Keep only up to max_consecutive occurrences
        if count > max_consecutive:
            cleaned.extend([current_word] * max_consecutive)
            print(f"  - Reduced '{current_word}' from {count} to {max_consecutive} occurrences")
        else:
            cleaned.extend([current_word] * count)
        
        i += count
    
    return ' '.join(cleaned)


def clean_text(text):
    """Run every cleaning pass over one block of text and return the result."""
    # Pass 1: Remove repetitive phrases (most important - catches longer patterns)
    text = remove_repetitive_phrases(text, min_phrase_words=3)

    # Pass 2: Remove excessive word repetitions
    text = remove_word_repetitions(text, max_consecutive=2)

    # Pass 3: Remove similar/related words (like "study studies studied")
    text = remove_similar_word_sequences(text)

    # Pass 4: Another pass for phrases (catches nested patterns)
    text = remove_repetitive_phrases(text, min_phrase_words=3)

    # Pass 5: Final cleanup of word repetitions. Keep two occurrences unless
    # AGGRESSIVE is set - English genuinely doubles words ("had had",
    # "that that", "New York, New York") and collapsing to one corrupts them.
    text = remove_word_repetitions(text, max_consecutive=1 if AGGRESSIVE else 2)

    # Collapse runs of spaces/tabs, but never newlines - see clean_transcript_file
    return re.sub(r'[^\S\n]+', ' ', text).strip()


def clean_transcript_file(file_path, write_backup=True, quiet=False):
    """Clean a single transcript file in place.

    Each line is cleaned independently so document structure survives - PPTX
    output carries headers like '### PowerPoint Slide Content ###' and
    '--- Slide 3 ---' that would otherwise be flattened into the prose.

    Set write_backup=False to skip the '<name>_backup.txt' copy.
    Returns True if the file was changed.
    """
    def say(msg):
        if not quiet:
            print(msg)

    say(f"\nProcessing: {file_path}")

    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            original_text = f.read()

        if not original_text.strip():
            say("  - File is empty, skipping")
            return False

        # Clean line by line so blank lines and headers keep their place
        cleaned_lines = [clean_text(line) if line.strip() else ""
                         for line in original_text.split("\n")]
        cleaned_text = "\n".join(cleaned_lines).strip()

        # Save if changes were made
        if cleaned_text != original_text:
            if write_backup:
                backup_path = str(file_path).replace('.txt', '_backup.txt')
                with open(backup_path, 'w', encoding='utf-8') as f:
                    f.write(original_text)
                say(f"  - Backup saved to: {backup_path}")

            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(cleaned_text)

            original_size = len(original_text)
            cleaned_size = len(cleaned_text)
            reduction = ((original_size - cleaned_size) / original_size) * 100

            say(f"  ✓ Cleaned successfully!")
            say(f"  - Original size: {original_size} chars")
            say(f"  - Cleaned size: {cleaned_size} chars")
            say(f"  - Reduction: {reduction:.1f}%")
            return True
        else:
            say("  - No repetitions found")
            return False

    except Exception as e:
        say(f"  ✗ Error processing file: {e}")
        return False


def clean_output_folder(folder_path='output'):
    """Clean all transcript files in the output folder."""
    folder = Path(folder_path)
    
    if not folder.exists():
        print(f"Error: Folder '{folder_path}' does not exist")
        return
    
    # Find all .txt files (excluding backups)
    txt_files = [f for f in folder.glob('*.txt') if not f.name.endswith('_backup.txt')]
    
    if not txt_files:
        print(f"No transcript files found in '{folder_path}'")
        return
    
    print(f"Found {len(txt_files)} transcript file(s) to process")
    print("=" * 60)
    
    for txt_file in txt_files:
        clean_transcript_file(txt_file)
    
    print("\n" + "=" * 60)
    print("✓ Processing complete!")


if __name__ == "__main__":
    clean_output_folder()
