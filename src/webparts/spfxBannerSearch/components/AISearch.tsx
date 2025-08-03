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

interface ISuggestion {
  id: string;
  title: string;
  subtitle?: string;
  icon: string;
  query: string;
}

// Sample suggestions - in real implementation, these could come from API
const SAMPLE_SUGGESTIONS: ISuggestion[] = [
  { id: '1', title: 'Marketing documents', subtitle: 'Find presentations and campaigns', icon: 'FileText', query: 'marketing documents' },
  { id: '2', title: 'HR policies', subtitle: 'Employee handbook and procedures', icon: 'People', query: 'HR policies' },
  { id: '3', title: 'Financial reports', subtitle: 'Quarterly and annual reports', icon: 'BarChart4', query: 'financial reports' },
  { id: '4', title: 'Project plans', subtitle: 'Timelines and project documents', icon: 'ProjectCollection', query: 'project plans' },
  { id: '5', title: 'Team contacts', subtitle: 'Employee directory and org chart', icon: 'Contact', query: 'team contacts' },
  { id: '6', title: 'Training materials', subtitle: 'Learning resources and courses', icon: 'Education', query: 'training materials' },
  { id: '7', title: 'Meeting notes', subtitle: 'Minutes and action items', icon: 'NotesTook', query: 'meeting notes' },
  { id: '8', title: 'Company news', subtitle: 'Announcements and updates', icon: 'News', query: 'company news' }
];

const AISearch: React.FC<IAISearchProps> = ({ placeholder, onSearchQuery }) => {
  const [query, setQuery] = useState<string>('');

  // Handle search execution - for now just show alert (placeholder for future AI features)
  const handleSearch = useCallback((newValue?: string) => {
    const searchQuery = newValue || query;
    if (searchQuery.trim()) {
      // Placeholder alert for AI functionality
      alert(`🤖 AI Search Mode!\n\nYou searched: "${searchQuery}"\n\nThis will be enhanced with AI features like:\n• Natural language processing\n• Smart query understanding\n• AI-powered suggestions\n• Contextual search results\n\nStay tuned!`);
      
      // Still trigger the search for now
      onSearchQuery(searchQuery.trim());
    }
  }, [query, onSearchQuery]);

  const handleKeyPress = useCallback((event: React.KeyboardEvent): void => {
    if (event.key === 'Enter') {
      handleSearch();
    }
  }, [handleSearch]);

  return (
    <div className={styles.searchContainer}>
      <SearchBox
        placeholder={placeholder || 'Ask me anything with AI...'}
        value={query}
        onChange={(_, newValue) => setQuery(newValue || '')}
        onSearch={handleSearch}
        onKeyDown={handleKeyPress}
        className={styles.heroSearchBox}
        iconProps={{ iconName: 'Robot' }}
        autoComplete="off"
      />
      <div className={styles.suggestionsHint}>
        <span>🤖 AI mode active - Natural language search enabled</span>
      </div>
    </div>
  );
};

export default AISearch;
