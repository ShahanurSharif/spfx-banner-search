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
import { useState, useCallback, useMemo } from 'react';
import styles from './SpfxBannerSearch.module.scss';
import type { ISpfxBannerSearchProps } from './ISpfxBannerSearchProps';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { ThemeProvider } from '@fluentui/react/lib/Theme';
import { Icon } from '@fluentui/react/lib/Icon';
import AISearch from './AISearch';
// ...existing code...
import { WebPartContext } from '@microsoft/sp-webpart-base';

import { SharePointSearchService } from '../../../services/SharePointSearchService';
import { useTypeahead } from '../../../hooks/useTypeahead';

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








// Enhanced search box component with SharePoint Search type-ahead suggestions (refactored)
const HeroSearchBox: React.FC<{
  placeholder: string;
  onSearch: (query: string) => void;
  enableSuggestions: boolean;
  semanticColors: Partial<import('@fluentui/react/lib/Styling').ISemanticColors>;
  context: WebPartContext;
  searchSiteUrl?: string;
  debugSuggestions?: boolean;
}> = React.memo(({ placeholder, onSearch, enableSuggestions, semanticColors, context, searchSiteUrl, debugSuggestions }) => {
  const service = useMemo(() => new SharePointSearchService(context, searchSiteUrl, debugSuggestions), [context, searchSiteUrl, debugSuggestions]);
  const {
    value: searchValue,
    onChange,
    suggestions,
    open: showSuggestions,
    loading: isSearching,
  // error, // not used
    setOpen: setShowSuggestions,
    setSuggestions
  } = useTypeahead(enableSuggestions ? service.fetchSuggestions.bind(service) : async () => [], 250);
  const [highlightedIndex, setHighlightedIndex] = useState<number>(-1);

  // Keyboard navigation
  const handleKeyDown = useCallback((event: React.KeyboardEvent) => {
    if (!showSuggestions || suggestions.length === 0) {
      if (event.key === 'Enter') {
        setShowSuggestions(false);
        setSuggestions([]);
        onSearch(searchValue);
      }
      return;
    }
    switch (event.key) {
      case 'ArrowDown':
        event.preventDefault();
        setHighlightedIndex(prev => prev < suggestions.length - 1 ? prev + 1 : 0);
        break;
      case 'ArrowUp':
        event.preventDefault();
        setHighlightedIndex(prev => prev > 0 ? prev - 1 : suggestions.length - 1);
        break;
      case 'Enter':
        event.preventDefault();
        if (highlightedIndex >= 0 && highlightedIndex < suggestions.length) {
          onChange(suggestions[highlightedIndex].suggestionTitle);
          setShowSuggestions(false);
          setSuggestions([]);
        } else {
          setShowSuggestions(false);
          setSuggestions([]);
          onSearch(searchValue);
        }
        break;
      case 'Escape':
        setShowSuggestions(false);
        setHighlightedIndex(-1);
        break;
    }
  }, [showSuggestions, suggestions, highlightedIndex, searchValue, onSearch, onChange, setShowSuggestions, setSuggestions]);

  // Handle input focus
  const handleFocus = useCallback((): void => {
    if (searchValue.trim() && suggestions.length > 0) {
      setShowSuggestions(true);
    }
  }, [searchValue, suggestions, setShowSuggestions]);

  return (
    <div className={styles.searchContainer}>
      <SearchBox
        placeholder={placeholder}
        value={searchValue}
        onChange={(_, newValue) => onChange(newValue || '')}
        onSearch={() => { setShowSuggestions(false); setSuggestions([]); onSearch(searchValue); }}
        onKeyDown={handleKeyDown}
        onFocus={handleFocus}
        className={styles.heroSearchBox}
        autoComplete="off"
        aria-expanded={showSuggestions}
        aria-haspopup="listbox"
        role="combobox"
      />
      {showSuggestions && suggestions.length > 0 && (
        <div
          className={styles.suggestionsDropdown}
          role="listbox"
          aria-label="Search results"
        >
          {isSearching && (
            <div className={styles.searchingIndicator}>Searching...</div>
          )}
          {suggestions.map((item, index) => (
            <div
              key={item.id}
              className={`${styles.suggestionItem} ${index === highlightedIndex ? styles.highlighted : ''}`}
              onClick={() => { onChange(item.suggestionTitle); setShowSuggestions(false); setSuggestions([]); }}
              role="option"
              aria-selected={index === highlightedIndex}
              onMouseEnter={() => setHighlightedIndex(index)}
            >
              <div className={styles.suggestionText}>
                <div className={styles.suggestionTitle}>{item.suggestionTitle}</div>
                <div className={styles.suggestionSubtitle}>{item.suggestionSubtitle}</div>
              </div>
            </div>
          ))}
        </div>
      )}
      {showSuggestions && suggestions.length === 0 && searchValue.trim() && !isSearching && (
        <div className={styles.suggestionsDropdown}>
          <div className={styles.noSuggestions}>
            No documents found. Press Enter to search for &quot;{searchValue}&quot;
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
    hasTeamsContext,
    context
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
                context={context}
                searchSiteUrl={props.searchSiteUrl}
                debugSuggestions={props.debugSuggestions}
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
