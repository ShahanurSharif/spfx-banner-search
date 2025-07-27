/**
 * AI Search Component
 * 
 * Temporary component for AI-powered search functionality.
 * Currently shows an alert when search is triggered.
 * Will be enhanced later with actual AI search capabilities.
 */

import * as React from 'react';
import { useState, useCallback } from 'react';
import { TextField } from '@fluentui/react/lib/TextField';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import styles from './SpfxBannerSearch.module.scss';

export interface IAISearchProps {
  placeholder: string;
  onSearchQuery: (queryText: string) => void;
}

const AISearch: React.FC<IAISearchProps> = ({ placeholder, onSearchQuery }) => {
  const [query, setQuery] = useState<string>('');

  const handleSearch = useCallback(() => {
    if (query.trim()) {
      // Temporary alert - will be replaced with actual AI search logic
      alert(`AI Search triggered with query: "${query}"`);
      onSearchQuery(query);
    }
  }, [query, onSearchQuery]);

  const handleKeyPress = useCallback((event: React.KeyboardEvent) => {
    if (event.key === 'Enter' && !event.shiftKey) {
      event.preventDefault();
      handleSearch();
    }
  }, [handleSearch]);

  return (
    <div className={styles.searchContainer}>
      <div className={styles.aiSearchWrapper}>
        <TextField
          placeholder={placeholder || 'Ask me anything...'}
          value={query}
          onChange={(_, newValue) => setQuery(newValue || '')}
          onKeyPress={handleKeyPress}
          multiline
          rows={3}
          resizable={false}
          styles={{
            root: {
              width: '100%',
              maxWidth: '600px'
            },
            field: {
              backgroundColor: 'rgba(255, 255, 255, 0.95)',
              border: 'none',
              borderRadius: '28px',
              fontSize: '1.125rem',
              padding: '20px 24px',
              minHeight: '100px',
              boxShadow: '0 8px 32px rgba(0, 0, 0, 0.12)',
              backdropFilter: 'blur(10px)',
              transition: 'all 0.3s cubic-bezier(0.4, 0, 0.2, 1)',
              lineHeight: '1.5',
              fontFamily: 'inherit',
              '&:hover': {
                backgroundColor: 'rgba(255, 255, 255, 0.98)',
                boxShadow: '0 12px 40px rgba(0, 0, 0, 0.15)',
                transform: 'translateY(-1px)'
              },
              '&:focus': {
                backgroundColor: 'rgba(255, 255, 255, 1)',
                boxShadow: '0 16px 48px rgba(0, 0, 0, 0.2)',
                transform: 'translateY(-2px)'
              },
              '&::placeholder': {
                color: '#605e5c',
                fontStyle: 'italic'
              }
            }
          }}
        />
        <PrimaryButton
          text="Ask AI"
          onClick={handleSearch}
          disabled={!query.trim()}
          className={styles.aiSearchButton}
        />
      </div>
    </div>
  );
};

export default AISearch;
