"""
Text processing using classical NLP techniques
No LLMs - only local, deterministic algorithms
"""

import re
from typing import List, Dict, Tuple
import numpy as np

# Optional imports - graceful degradation if not available
try:
    from sklearn.feature_extraction.text import TfidfVectorizer
    SKLEARN_AVAILABLE = True
except ImportError:
    SKLEARN_AVAILABLE = False

try:
    import nltk
    from nltk.tokenize import sent_tokenize, word_tokenize
    from nltk.corpus import stopwords
    
    # Download required NLTK data
    try:
        nltk.data.find('tokenizers/punkt')
    except LookupError:
        nltk.download('punkt', quiet=True)
    
    try:
        nltk.data.find('corpora/stopwords')
    except LookupError:
        nltk.download('stopwords', quiet=True)
    
    NLTK_AVAILABLE = True
except ImportError:
    NLTK_AVAILABLE = False

from .error_handler import ErrorHandler


class TextProcessor:
    """
    Classical NLP text processing
    Uses only local, deterministic algorithms
    """
    
    def __init__(self):
        self.error_handler = ErrorHandler()
        self.vectorizer = None
        
        if SKLEARN_AVAILABLE:
            self.vectorizer = TfidfVectorizer(
                max_features=100,
                stop_words='english',
                ngram_range=(1, 2)
            )
        
        if not NLTK_AVAILABLE:
            self.error_handler.log_warning(
                "NLTK not available. Using simple text processing."
            )
    
    def summarize_to_bullets(
        self, 
        text: str, 
        max_bullets: int = 5,
        min_sentence_length: int = 20
    ) -> List[str]:
        """
        Extract most important sentences as bullet points
        Uses TextRank algorithm (local, deterministic)
        
        Args:
            text: Input text to summarize
            max_bullets: Maximum number of bullets to return
            min_sentence_length: Minimum characters per sentence
            
        Returns:
            List of bullet point strings
        """
        if not text or not text.strip():
            return []
        
        # Tokenize into sentences
        if NLTK_AVAILABLE:
            sentences = sent_tokenize(text)
        else:
            sentences = self._simple_sentence_split(text)
        
        # Filter short sentences
        sentences = [s for s in sentences if len(s) >= min_sentence_length]
        
        if len(sentences) <= max_bullets:
            return [self._clean_bullet(s) for s in sentences]
        
        # Use TF-IDF scoring if available
        if SKLEARN_AVAILABLE and self.vectorizer:
            try:
                scored_sentences = self._score_sentences_tfidf(sentences)
                top_sentences = self._get_top_sentences(
                    scored_sentences, 
                    max_bullets
                )
                return [self._clean_bullet(s) for s in top_sentences]
            except Exception as e:
                self.error_handler.log_warning(f"TF-IDF scoring failed: {e}")
        
        # Fallback: return first N sentences
        return [self._clean_bullet(s) for s in sentences[:max_bullets]]
    
    def _score_sentences_tfidf(self, sentences: List[str]) -> List[Tuple[str, float]]:
        """Score sentences using TF-IDF"""
        try:
            tfidf_matrix = self.vectorizer.fit_transform(sentences)
            scores = np.array(tfidf_matrix.sum(axis=1)).flatten()
            
            return list(zip(sentences, scores))
        except:
            # Fallback to equal scores
            return list(zip(sentences, [1.0] * len(sentences)))
    
    def _get_top_sentences(
        self, 
        scored_sentences: List[Tuple[str, float]], 
        n: int
    ) -> List[str]:
        """Get top N sentences maintaining original order"""
        # Sort by score
        sorted_sentences = sorted(
            scored_sentences, 
            key=lambda x: x[1], 
            reverse=True
        )
        
        # Get top N
        top_n = sorted_sentences[:n]
        
        # Restore original order
        sentence_to_score = {sent: score for sent, score in scored_sentences}
        top_sentences = [sent for sent, _ in top_n]
        top_sentences.sort(
            key=lambda x: [s for s, _ in scored_sentences].index(x)
        )
        
        return top_sentences
    
    def extract_keywords(self, text: str, top_n: int = 10) -> List[str]:
        """
        Extract key terms using TF-IDF
        
        Args:
            text: Input text
            top_n: Number of keywords to extract
            
        Returns:
            List of keyword strings
        """
        if not SKLEARN_AVAILABLE or not self.vectorizer:
            return self._simple_keyword_extract(text, top_n)
        
        try:
            tfidf_matrix = self.vectorizer.fit_transform([text])
            feature_names = self.vectorizer.get_feature_names_out()
            scores = tfidf_matrix.toarray()[0]
            
            # Get top N keywords
            top_indices = scores.argsort()[-top_n:][::-1]
            keywords = [feature_names[i] for i in top_indices]
            
            return keywords
        except Exception as e:
            self.error_handler.log_warning(f"Keyword extraction failed: {e}")
            return self._simple_keyword_extract(text, top_n)
    
    def _simple_keyword_extract(self, text: str, top_n: int) -> List[str]:
        """Simple frequency-based keyword extraction"""
        # Remove punctuation and lowercase
        words = re.findall(r'\b[a-z]{4,}\b', text.lower())
        
        # Count frequencies
        word_freq = {}
        for word in words:
            if len(word) >= 4:  # Skip short words
                word_freq[word] = word_freq.get(word, 0) + 1
        
        # Get top N
        sorted_words = sorted(
            word_freq.items(), 
            key=lambda x: x[1], 
            reverse=True
        )
        
        return [word for word, _ in sorted_words[:top_n]]
    
    def truncate_smart(self, text: str, max_length: int) -> str:
        """
        Truncate text at sentence boundary
        
        Args:
            text: Text to truncate
            max_length: Maximum character length
            
        Returns:
            Truncated text
        """
        if len(text) <= max_length:
            return text
        
        # Try to split at sentence boundary
        if NLTK_AVAILABLE:
            sentences = sent_tokenize(text)
        else:
            sentences = self._simple_sentence_split(text)
        
        result = ""
        for sent in sentences:
            if len(result) + len(sent) + 1 <= max_length:
                result += sent + " "
            else:
                break
        
        result = result.strip()
        
        # Add ellipsis if truncated
        if result and len(result) < len(text):
            # Remove incomplete last sentence if any
            if not result.endswith(('.', '!', '?')):
                # Find last complete sentence
                last_end = max(
                    result.rfind('.'),
                    result.rfind('!'),
                    result.rfind('?')
                )
                if last_end > 0:
                    result = result[:last_end + 1]
            result += "..."
        
        return result if result else text[:max_length] + "..."
    
    def _simple_sentence_split(self, text: str) -> List[str]:
        """Simple sentence splitting without NLTK"""
        # Split on period, exclamation, question mark
        sentences = re.split(r'[.!?]+', text)
        sentences = [s.strip() for s in sentences if s.strip()]
        return sentences
    
    def _clean_bullet(self, text: str) -> str:
        """Clean and format bullet point text"""
        text = text.strip()
        
        # Remove multiple spaces
        text = re.sub(r'\s+', ' ', text)
        
        # Capitalize first letter
        if text and text[0].islower():
            text = text[0].upper() + text[1:]
        
        # Ensure ends with period if not already punctuated
        if text and text[-1] not in '.!?':
            text += '.'
        
        return text
    
    def format_bullet_list(self, items: List[str]) -> List[str]:
        """Format list of items as proper bullet points"""
        formatted = []
        
        for item in items:
            cleaned = self._clean_bullet(item)
            formatted.append(cleaned)
        
        return formatted
    
    def split_long_content(
        self, 
        content: List[str], 
        max_items: int = 6,
        max_chars_per_item: int = 200
    ) -> List[List[str]]:
        """
        Split content into multiple slides if too long
        
        Args:
            content: List of content items
            max_items: Maximum items per slide
            max_chars_per_item: Maximum characters per item
            
        Returns:
            List of content lists (one per slide)
        """
        if len(content) <= max_items:
            return [content]
        
        slides = []
        current_slide = []
        
        for item in content:
            # Truncate individual items if too long
            if len(item) > max_chars_per_item:
                item = self.truncate_smart(item, max_chars_per_item)
            
            current_slide.append(item)
            
            if len(current_slide) >= max_items:
                slides.append(current_slide)
                current_slide = []
        
        # Add remaining items
        if current_slide:
            slides.append(current_slide)
        
        return slides
    
    def rank_content_by_relevance(
        self, 
        items: List[str], 
        query: str
    ) -> List[str]:
        """
        Rank content items by relevance to query
        
        Args:
            items: List of content items
            query: Query string to rank against
            
        Returns:
            Sorted list of items (most relevant first)
        """
        if not SKLEARN_AVAILABLE or not items:
            return items
        
        try:
            # Create corpus with query + items
            corpus = [query] + items
            
            # Fit TF-IDF
            tfidf_matrix = self.vectorizer.fit_transform(corpus)
            
            # Calculate similarity to query (first item)
            query_vector = tfidf_matrix[0:1]
            item_vectors = tfidf_matrix[1:]
            
            # Cosine similarity
            similarities = (item_vectors * query_vector.T).toarray().flatten()
            
            # Sort items by similarity
            ranked_items = [
                item for _, item in sorted(
                    zip(similarities, items), 
                    reverse=True
                )
            ]
            
            return ranked_items
            
        except Exception as e:
            self.error_handler.log_warning(f"Content ranking failed: {e}")
            return items
