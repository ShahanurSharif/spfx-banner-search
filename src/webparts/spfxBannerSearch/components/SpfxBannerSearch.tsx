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
// import { SearchBox } from '@fluentui/react/lib/SearchBox'; // Temporarily commented out for testing
import { ThemeProvider } from '@fluentui/react/lib/Theme';
import { Icon } from '@fluentui/react/lib/Icon';

// Helper function to get file type icon
const getFileTypeIcon = (fileType: string): string => {
  const lowerType = fileType?.toLowerCase() || '';
  switch (lowerType) {
    case 'pdf': return 'PDF';
    case 'doc':
    case 'docx': return 'WordDocument';
    case 'xls':
    case 'xlsx': return 'ExcelDocument';
    case 'ppt':
    case 'pptx': return 'PowerPointDocument';
    case 'txt': return 'TextDocument';
    case 'html':
    case 'htm': return 'FileHTML';
    default: return 'Page';
  }
};
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
  console.debug("[HeroSearchBox] Component is rendering with props:", { placeholder, enableSuggestions });
  const service = useMemo(() => new SharePointSearchService(context, searchSiteUrl, debugSuggestions), [context, searchSiteUrl, debugSuggestions]);
  
  // Create a stable fetchFn to prevent infinite loops
  const fetchFn = useCallback(
    (term: string, signal?: AbortSignal) => {
      if (!enableSuggestions) return Promise.resolve([]);
      return service.fetchSuggestions(term, signal);
    },
    [service, enableSuggestions]
  );
  
  const {
    value: searchValue,
    onChange,
    suggestions,
    open: suggestionsOpen,
    loading: isSearching,
  // error, // not used
    setOpen: setShowSuggestions,
    setSuggestions
  } = useTypeahead(fetchFn, 250);
  
  // Use the open state from useTypeahead hook
  const showSuggestions = suggestionsOpen;
  
  // Debug logging (only when enabled)
  if (debugSuggestions) {
    console.debug("[HeroSearchBox] showSuggestions:", showSuggestions, "suggestions.length:", suggestions.length);
  }
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

  // Simple focus handler - let useTypeahead manage the suggestions
  const handleFocus = useCallback((): void => {
    // Focus handling is now managed by useTypeahead hook
  }, []);

  // Handle input blur with delay to allow for clicks
  const handleBlur = useCallback((event: React.FocusEvent) => {
    // Delay hiding suggestions to allow for suggestion clicks
    setTimeout(() => {
      if (event.currentTarget && !event.currentTarget.contains(document.activeElement)) {
        setShowSuggestions(false);
        setHighlightedIndex(-1);
      }
    }, 150);
  }, [setShowSuggestions]);

  console.debug("[HeroSearchBox] Rendering SearchBox with value:", searchValue);
  
  return (
    <div className={styles.searchContainer}>
      {/* Temporary: Using regular input to test if SearchBox is the issue */}
      <input
        type="text"
        placeholder={placeholder}
        value={searchValue}
        onChange={(e) => {
          console.debug("[Input] onChange called with:", e.target.value);
          onChange(e.target.value || '');
        }}
        onKeyDown={handleKeyDown}
        onFocus={handleFocus}
        onBlur={handleBlur}
        className={styles.heroSearchBox}
        autoComplete="off"
        aria-expanded={showSuggestions}
        aria-haspopup="listbox"
        role="combobox"
        style={{
          width: '100%',
          border: 'none',
          fontSize: '1rem',
          padding: '0 20px',
          borderRadius: '4px',
          outline: 'none'
        }}
      />
      {searchValue.trim().length > 0 && (
        <div
          className={styles.suggestionsDropdown}
          role="listbox"
          aria-label="Search results"
          style={{
            position: 'absolute',
            top: '100%',
            left: 0,
            right: 0,
            background: 'white',
            border: '1px solid #ccc',
            borderRadius: '4px',
            boxShadow: '0 4px 8px rgba(0,0,0,0.1)',
            zIndex: 1000,
            maxHeight: '300px',
            overflowY: 'auto'
          }}
        >
          {isSearching && (
            <div className={styles.searchingIndicator} style={{ padding: '16px 20px', textAlign: 'center', color: '#666' }}>
              Searching...
            </div>
          )}
          {!isSearching && suggestions.length === 0 && (
            <div style={{ padding: '16px 20px', textAlign: 'center', color: '#666', fontStyle: 'italic' }}>
              No content found for &quot;{searchValue}&quot;
            </div>
          )}
          {!isSearching && suggestions.length > 0 && suggestions.map((item, index) => (
            <div
              key={item.id}
              className={`${styles.suggestionItem} ${index === highlightedIndex ? styles.highlighted : ''}`}
              onClick={() => { 
                // If item has a path, open it in new tab, otherwise perform search
                if (item.path && item.path.trim()) {
                  window.open(item.path, '_blank', 'noopener,noreferrer');
                } else {
                  onChange(item.suggestionTitle); 
                  onSearch(item.suggestionTitle);
                }
                setShowSuggestions(false); 
                setSuggestions([]);
              }}
              role="option"
              aria-selected={index === highlightedIndex}
              onMouseEnter={() => setHighlightedIndex(index)}
              style={{
                padding: '12px 20px',
                cursor: 'pointer',
                borderBottom: '1px solid #eee',
                backgroundColor: index === highlightedIndex ? '#f0f0f0' : 'white'
              }}
            >
              <Icon
                iconName={getFileTypeIcon(item.fileType || '')}
                className={styles.fileIcon}
              />
              <div className={styles.suggestionText} style={{ color: '#333' }}>
                <div className={styles.suggestionTitle} style={{ fontWeight: 'bold', marginBottom: '4px' }}>
                  {item.suggestionTitle}
                </div>
                <div className={styles.suggestionSubtitle} style={{ fontSize: '12px', color: '#666' }}>
                  {item.suggestionSubtitle}
                </div>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
});


// Main hero banner component
const SpfxBannerSearch: React.FC<ISpfxBannerSearchProps> = (props) => {
  console.debug("[SpfxBannerSearch] Main component is rendering");
  const {
    gradientStartColor,
    gradientEndColor,
    showCircleAnimation,
    minHeight,
    titleFontSize,
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
    '--title-font-size': `${titleFontSize}px`,
    '--body-text': semanticColors?.bodyText || (isDarkTheme ? '#ffffff' : '#323130'),
    '--link-color': semanticColors?.link || '#0078d4',
    '--link-hover': semanticColors?.linkHovered || '#106ebe'
  } as React.CSSProperties), [gradientStartColor, gradientEndColor, minHeight, titleFontSize, semanticColors, isDarkTheme]);

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

console.debug("[SpfxBannerSearch] Module loaded");

export default SpfxBannerSearch;
