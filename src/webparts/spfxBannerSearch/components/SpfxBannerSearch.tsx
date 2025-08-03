/**
 * SPFx Hero Banner Search Web Part
 * 
 * Full-bleed hero banner component with centralized search functionality.
 * Integrates with PnP Modern Search through dynamic data publishing.
 * 
 * Features:
 * - Full-width responsive design (supports supportsFullBleed: true)
 * - Configurable gradient background colors
 * - Optional floating circle animations
 * - Centralized search box with suggestions support
 * - Dynamic data publishing for search queries
 * - Theme-aware styling with CSS variables
 * - Accessibility and responsive design
 * - Teams context support
 * 
 * Dynamic Data Integration:
 * - Publishes 'inputQueryText' property for connection to Search Results web parts
 * - Compatible with PnP Modern Search Results web part connections
 * 
 * Performance Optimizations:
 * - React.memo for child components
 * - useMemo for expensive calculations
 * - useCallback for event handlers
 * - CSS animations with hardware acceleration
 * 
 * @author Generated for SPFx Hero Banner Search
 */

import * as React from 'react';
import { useState, useCallback, useMemo, useRef, useEffect } from 'react';
import styles from './SpfxBannerSearch.module.scss';
import type { ISpfxBannerSearchProps } from './ISpfxBannerSearchProps';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { ThemeProvider } from '@fluentui/react/lib/Theme';
import { Icon } from '@fluentui/react/lib/Icon';
import AISearch from './AISearch';

// Animation component for floating circles
const AnimatedCircles: React.FC<{ show: boolean }> = React.memo(({ show }) => {
  if (!show) return null;
  
  return (
    <div className={styles.animationLayer}>
      {[...Array(6)].map((_, i) => (
        <div
          key={i}
          className={`${styles.floatingCircle} ${styles[`circle${i + 1}`]}`}
          style={{
            animationDelay: `${i * 2}s`,
            animationDuration: `${8 + i * 2}s`
          }}
        />
      ))}
    </div>
  );
});

// AI Search Toggle component
const AIToggle: React.FC<{ 
  isActive: boolean; 
  onToggle: () => void;
}> = React.memo(({ isActive, onToggle }) => {
  return (
    <div className={styles.aiToggleContainer}>
      <button
        className={`${styles.aiToggleButton} ${isActive ? styles.active : ''}`}
        onClick={onToggle}
        title={isActive ? "Switch to regular search" : "Switch to AI search"}
        aria-label={isActive ? "Disable AI search" : "Enable AI search"}
        type="button"
      >
        <Icon 
          iconName={isActive ? "Robot" : "Lightbulb"} 
          className={styles.aiToggleIcon}
        />
        <span className={styles.aiToggleText}>
          {isActive ? "AI ON" : "AI"}
        </span>
      </button>
    </div>
  );
});

// Suggestion interface for regular search
interface ISuggestion {
  id: string;
  title: string;
  subtitle?: string;
  icon: string;
  query: string;
}

// Sample suggestions for regular search
const SEARCH_SUGGESTIONS: ISuggestion[] = [
  { id: '1', title: 'Marketing documents', subtitle: 'Find presentations and campaigns', icon: 'FileText', query: 'marketing documents' },
  { id: '2', title: 'HR policies', subtitle: 'Employee handbook and procedures', icon: 'People', query: 'HR policies' },
  { id: '3', title: 'Financial reports', subtitle: 'Quarterly and annual reports', icon: 'BarChart4', query: 'financial reports' },
  { id: '4', title: 'Project plans', subtitle: 'Timelines and project documents', icon: 'ProjectCollection', query: 'project plans' },
  { id: '5', title: 'Team contacts', subtitle: 'Employee directory and org chart', icon: 'Contact', query: 'team contacts' },
  { id: '6', title: 'Training materials', subtitle: 'Learning resources and courses', icon: 'Education', query: 'training materials' },
  { id: '7', title: 'Meeting notes', subtitle: 'Minutes and action items', icon: 'NotesTook', query: 'meeting notes' },
  { id: '8', title: 'Company news', subtitle: 'Announcements and updates', icon: 'News', query: 'company news' }
];

// Enhanced search box component with suggestions
const HeroSearchBox: React.FC<{
  placeholder: string;
  onSearch: (query: string) => void;
  enableSuggestions: boolean;
  semanticColors: Partial<import('@fluentui/react/lib/Styling').ISemanticColors>;
}> = React.memo(({ placeholder, onSearch, enableSuggestions, semanticColors }) => {
  const [searchValue, setSearchValue] = useState('');
  const [showSuggestions, setShowSuggestions] = useState<boolean>(false);
  const [filteredSuggestions, setFilteredSuggestions] = useState<ISuggestion[]>([]);
  const [highlightedIndex, setHighlightedIndex] = useState<number>(-1);
  
  const searchBoxRef = useRef<HTMLDivElement>(null);
  const suggestionsRef = useRef<HTMLDivElement>(null);

  // Filter suggestions based on query
  useEffect(() => {
    if (!searchValue.trim() || !enableSuggestions) {
      setFilteredSuggestions([]);
      setShowSuggestions(false);
      return;
    }

    const filtered = SEARCH_SUGGESTIONS.filter(suggestion =>
      suggestion.title.toLowerCase().includes(searchValue.toLowerCase()) ||
      suggestion.subtitle?.toLowerCase().includes(searchValue.toLowerCase())
    );

    setFilteredSuggestions(filtered);
    setShowSuggestions(filtered.length > 0);
    setHighlightedIndex(-1);
  }, [searchValue, enableSuggestions]);

  // Handle search execution
  const executeSearch = useCallback((searchQuery: string) => {
    if (searchQuery.trim()) {
      setShowSuggestions(false);
      setSearchValue(searchQuery);
      onSearch(searchQuery.trim());
    }
  }, [onSearch]);

  // Handle suggestion selection
  const selectSuggestion = useCallback((suggestion: ISuggestion) => {
    executeSearch(suggestion.query);
  }, [executeSearch]);

  // Handle keyboard navigation
  const handleKeyDown = useCallback((event: React.KeyboardEvent) => {
    if (!showSuggestions || filteredSuggestions.length === 0) {
      if (event.key === 'Enter') {
        executeSearch(searchValue);
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
          executeSearch(searchValue);
        }
        break;
      
      case 'Escape':
        setShowSuggestions(false);
        setHighlightedIndex(-1);
        break;
    }
  }, [showSuggestions, filteredSuggestions, highlightedIndex, searchValue, executeSearch, selectSuggestion]);

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
    if (searchValue.trim() && filteredSuggestions.length > 0) {
      setShowSuggestions(true);
    }
  }, [searchValue, filteredSuggestions]);

  return (
    <div className={styles.searchContainer} ref={searchBoxRef}>
      <SearchBox
        placeholder={placeholder}
        value={searchValue}
        onChange={(_, newValue) => setSearchValue(newValue || '')}
        onSearch={executeSearch}
        onKeyDown={handleKeyDown}
        onFocus={handleFocus}
        className={styles.heroSearchBox}
        iconProps={{ iconName: 'Search' }}
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
          aria-label="Search suggestions"
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
      
      {showSuggestions && filteredSuggestions.length === 0 && searchValue.trim() && (
        <div className={styles.suggestionsDropdown} ref={suggestionsRef}>
          <div className={styles.noSuggestions}>
            No suggestions found. Press Enter to search for &quot;{searchValue}&quot;
          </div>
        </div>
      )}
    </div>
  );
});

// Main hero banner component
const SpfxBannerSearch: React.FC<ISpfxBannerSearchProps> = (props) => {
  const {
    gradientStartColor,
    gradientEndColor,
    showCircleAnimation,
    minHeight,
    bannerTitle,
    searchBoxPlaceholder,
    enableSuggestions,
    isDarkTheme,
    semanticColors,
    onSearchQuery,
    hasTeamsContext
  } = props;

  // Local state for AI search toggle (defaults to false, users can toggle)
  const [isAISearchActive, setIsAISearchActive] = useState<boolean>(false);

  // Memoized styles for performance
  const bannerStyles = useMemo(() => ({
    '--gradient-start': gradientStartColor,
    '--gradient-end': gradientEndColor,
    '--min-height': `${minHeight}px`,
    '--body-text': semanticColors?.bodyText || (isDarkTheme ? '#ffffff' : '#323130'),
    '--link-color': semanticColors?.link || '#0078d4',
    '--link-hover': semanticColors?.linkHovered || '#106ebe'
  } as React.CSSProperties), [gradientStartColor, gradientEndColor, minHeight, semanticColors, isDarkTheme]);

  // Theme provider configuration for Fluent UI components
  const theme = useMemo(() => ({
    palette: {
      themePrimary: gradientStartColor,
      themeLighterAlt: gradientEndColor,
      neutralPrimary: semanticColors?.bodyText || (isDarkTheme ? '#ffffff' : '#323130'),
      neutralSecondary: semanticColors?.bodySubtext || (isDarkTheme ? '#d0d0d0' : '#605e5c'),
      white: isDarkTheme ? '#1f1f1f' : '#ffffff',
      neutralLight: isDarkTheme ? '#2d2d2d' : '#f3f2f1'
    }
  }), [gradientStartColor, gradientEndColor, semanticColors, isDarkTheme]);

  // AI toggle handler
  const handleAIToggle = useCallback(() => {
    setIsAISearchActive(prev => !prev);
  }, []);

  // Optimized search handler
  const handleSearch = useCallback((queryText: string) => {
    onSearchQuery(queryText);
  }, [onSearchQuery]);

  // CSS classes for responsive design
  const containerClasses = `${styles.heroBanner} ${hasTeamsContext ? styles.teams : ''}`;

  return (
    <ThemeProvider theme={theme}>
      <section 
        className={containerClasses}
        style={bannerStyles}
        role="search"
        aria-label="Hero search banner"
      >
        {/* Animated background circles */}
        <AnimatedCircles show={showCircleAnimation} />
        
        {/* AI Search Toggle - Top Right Corner */}
        <AIToggle 
          isActive={isAISearchActive} 
          onToggle={handleAIToggle}
        />
        
        {/* Main content container */}
        <div className={styles.heroContent}>
          <div className={styles.searchWrapper}>
            <h1 className={styles.heroTitle}>
              {bannerTitle || 'Find What You Need'}
            </h1>
            
            {isAISearchActive ? (
              <AISearch
                placeholder={searchBoxPlaceholder || 'Ask me anything...'}
                onSearchQuery={handleSearch}
                enableSuggestions={enableSuggestions}
              />
            ) : (
              <HeroSearchBox
                placeholder={searchBoxPlaceholder}
                onSearch={handleSearch}
                enableSuggestions={enableSuggestions}
                semanticColors={semanticColors}
              />
            )}
          </div>
        </div>
        
        {/* Accessibility landmark */}
        <div className={styles.srOnly}>
          Full-width search interface for finding content
        </div>
      </section>
    </ThemeProvider>
  );
};

export default SpfxBannerSearch;
