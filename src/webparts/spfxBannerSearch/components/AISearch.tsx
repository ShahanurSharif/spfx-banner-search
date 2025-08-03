/**
 * Enhanced AI Search Component with Suggestions
 * 
 * Features:
 * - Real-time search suggestions dropdown
 * - Keyboard navigation (arrow keys, enter, escape)
 * - Smart suggestion filtering
 * - Proper search flow: type → suggestions → enter → results
 * - Click outside to close suggestions
 * - Accessible design with ARIA support
 */

import * as React from 'react';
import { useState, useCallback, useEffect, useRef } from 'react';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './SpfxBannerSearch.module.scss';

export interface IAISearchProps {
  placeholder: string;
  onSearchQuery: (queryText: string) => void;
  enableSuggestions?: boolean;
}

// Suggestion interface for AI search
interface ISuggestion {
  id: string;
  title: string;
  subtitle?: string;
  icon: string;
  query: string;
}

// AI-specific suggestions - different from regular search
const AI_SUGGESTIONS: ISuggestion[] = [
  { id: '1', title: 'Marketing documents', subtitle: 'Find presentations and campaigns', icon: 'FileText', query: 'marketing documents' },
  { id: '2', title: 'HR policies', subtitle: 'Employee handbook and procedures', icon: 'People', query: 'HR policies' },
  { id: '3', title: 'Financial reports', subtitle: 'Quarterly and annual reports', icon: 'BarChart4', query: 'financial reports' },
  { id: '4', title: 'Project plans', subtitle: 'Timelines and project documents', icon: 'ProjectCollection', query: 'project plans' },
  { id: '5', title: 'Team contacts', subtitle: 'Employee directory and org chart', icon: 'Contact', query: 'team contacts' },
  { id: '6', title: 'Training materials', subtitle: 'Learning resources and courses', icon: 'Education', query: 'training materials' },
  { id: '7', title: 'Meeting notes', subtitle: 'Minutes and action items', icon: 'NotesTook', query: 'meeting notes' },
  { id: '8', title: 'Company news', subtitle: 'Announcements and updates', icon: 'News', query: 'company news' }
];

const AISearch: React.FC<IAISearchProps> = ({ placeholder, onSearchQuery, enableSuggestions = true }) => {
  const [query, setQuery] = useState<string>('');
  const [showSuggestions, setShowSuggestions] = useState<boolean>(false);
  const [filteredSuggestions, setFilteredSuggestions] = useState<ISuggestion[]>([]);
  const [highlightedIndex, setHighlightedIndex] = useState<number>(-1);
  
  const searchBoxRef = useRef<HTMLDivElement>(null);
  const suggestionsRef = useRef<HTMLDivElement>(null);

  // Filter suggestions based on query
  useEffect(() => {
    if (!query.trim() || !enableSuggestions) {
      setFilteredSuggestions([]);
      setShowSuggestions(false);
      return;
    }

    const filtered = AI_SUGGESTIONS.filter(suggestion =>
      suggestion.title.toLowerCase().includes(query.toLowerCase()) ||
      suggestion.subtitle?.toLowerCase().includes(query.toLowerCase())
    );

    setFilteredSuggestions(filtered);
    setShowSuggestions(filtered.length > 0);
    setHighlightedIndex(-1);
  }, [query, enableSuggestions]);

  // Handle search execution - shows AI placeholder alert
  const executeSearch = useCallback((searchQuery: string) => {
    if (searchQuery.trim()) {
      setShowSuggestions(false);
      setQuery(searchQuery);
      
      // AI placeholder alert
      alert(`🤖 AI Search Mode!\n\nYou searched: "${searchQuery}"\n\nThis will be enhanced with AI features like:\n• Natural language processing\n• Smart query understanding\n• AI-powered suggestions\n• Contextual search results\n\nStay tuned!`);
      
      // Still trigger the search for now
      onSearchQuery(searchQuery);
    }
  }, [onSearchQuery]);

  // Handle suggestion selection
  const selectSuggestion = useCallback((suggestion: ISuggestion) => {
    executeSearch(suggestion.query);
  }, [executeSearch]);

  // Handle keyboard navigation
  const handleKeyDown = useCallback((event: React.KeyboardEvent) => {
    if (!showSuggestions || filteredSuggestions.length === 0) {
      if (event.key === 'Enter') {
        executeSearch(query);
      }
      return;
    }

    switch (event.key) {
      case 'ArrowDown':
        event.preventDefault();
        setHighlightedIndex(prev => 
          prev < filteredSuggestions.length - 1 ? prev + 1 : 0
        );
        break;
      
      case 'ArrowUp':
        event.preventDefault();
        setHighlightedIndex(prev => 
          prev > 0 ? prev - 1 : filteredSuggestions.length - 1
        );
        break;
      
      case 'Enter':
        event.preventDefault();
        if (highlightedIndex >= 0 && highlightedIndex < filteredSuggestions.length) {
          selectSuggestion(filteredSuggestions[highlightedIndex]);
        } else {
          executeSearch(query);
        }
        break;
      
      case 'Escape':
        setShowSuggestions(false);
        setHighlightedIndex(-1);
        break;
    }
  }, [showSuggestions, filteredSuggestions, highlightedIndex, query, executeSearch, selectSuggestion]);

  // Handle click outside to close suggestions
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent): void => {
      if (searchBoxRef.current && !searchBoxRef.current.contains(event.target as Node) &&
          suggestionsRef.current && !suggestionsRef.current.contains(event.target as Node)) {
        setShowSuggestions(false);
      }
    };

    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  // Handle input focus
  const handleFocus = useCallback((): void => {
    if (query.trim() && filteredSuggestions.length > 0) {
      setShowSuggestions(true);
    }
  }, [query, filteredSuggestions]);

  return (
    <div className={styles.searchContainer} ref={searchBoxRef}>
      <SearchBox
        placeholder={placeholder || 'Ask me anything with AI...'}
        value={query}
        onChange={(_, newValue) => setQuery(newValue || '')}
        onSearch={executeSearch}
        onKeyDown={handleKeyDown}
        onFocus={handleFocus}
        className={styles.heroSearchBox}
        iconProps={{ iconName: 'Robot' }}
        autoComplete="off"
        aria-expanded={showSuggestions}
        aria-haspopup="listbox"
        role="combobox"
      />
      
      {showSuggestions && filteredSuggestions.length > 0 && (
        <div 
          className={styles.suggestionsDropdown}
          ref={suggestionsRef}
          role="listbox"
          aria-label="AI search suggestions"
        >
          {filteredSuggestions.map((suggestion, index) => (
            <div
              key={suggestion.id}
              className={`${styles.suggestionItem} ${index === highlightedIndex ? styles.highlighted : ''}`}
              onClick={() => selectSuggestion(suggestion)}
              role="option"
              aria-selected={index === highlightedIndex}
              onMouseEnter={() => setHighlightedIndex(index)}
            >
              <Icon iconName={suggestion.icon} className={styles.suggestionIcon} />
              <div className={styles.suggestionText}>
                <div className={styles.suggestionTitle}>{suggestion.title}</div>
                {suggestion.subtitle && (
                  <div className={styles.suggestionSubtitle}>{suggestion.subtitle}</div>
                )}
              </div>
            </div>
          ))}
        </div>
      )}
      
      {showSuggestions && filteredSuggestions.length === 0 && query.trim() && (
        <div className={styles.suggestionsDropdown} ref={suggestionsRef}>
          <div className={styles.noSuggestions}>
            No suggestions found. Press Enter to search for &quot;{query}&quot;
          </div>
        </div>
      )}
      
      <div className={styles.suggestionsHint}>
        <span>🤖 AI mode active - Natural language search enabled</span>
      </div>
    </div>
  );
};

export default AISearch;
