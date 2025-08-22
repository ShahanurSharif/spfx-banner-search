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
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

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

// SharePoint search result interface
interface ISharePointResult {
  id: string;
  title: string;
  subtitle?: string;
  url: string;
  fileType: string;
  lastModified: string;
  author?: string;
}



// Enhanced search box component with SharePoint search suggestions
const HeroSearchBox: React.FC<{
  placeholder: string;
  onSearch: (query: string) => void;
  enableSuggestions: boolean;
  semanticColors: Partial<import('@fluentui/react/lib/Styling').ISemanticColors>;
  context: WebPartContext;
}> = React.memo(({ placeholder, onSearch, enableSuggestions, semanticColors, context }) => {
  const [searchValue, setSearchValue] = useState('');
  const [showSuggestions, setShowSuggestions] = useState<boolean>(false);
  const [searchResults, setSearchResults] = useState<ISharePointResult[]>([]);
  const [highlightedIndex, setHighlightedIndex] = useState<number>(-1);
  const [isSearching, setIsSearching] = useState<boolean>(false);
  
  const searchBoxRef = useRef<HTMLDivElement>(null);
  const suggestionsRef = useRef<HTMLDivElement>(null);

  // Search other common document libraries
  const searchOtherLibraries = useCallback(async (query: string) => {
    if (!context) return;
    
    try {
      // Try searching in other common libraries
      const libraries = ['Shared Documents', 'Site Assets', 'Site Pages'];
      
      for (const libraryName of libraries) {
        try {
          const libraryUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(libraryName)}')/items?$filter=substringof('${encodeURIComponent(query)}',Title)&$select=Title,FileRef,FileLeafRef,Modified,Author/Title,File_x0020_Type&$expand=Author&$top=5&$orderby=Modified desc`;
          
          const response: SPHttpClientResponse = await context.spHttpClient.get(
            libraryUrl,
            SPHttpClient.configurations.v1
          );
          
          if (response.ok) {
            const data = await response.json();
            
            if (data.value && data.value.length > 0) {
              const results: ISharePointResult[] = data.value.map((item: any, index: number) => ({
                id: `${libraryName}-${index}`,
                title: item.Title || item.FileLeafRef || 'Untitled',
                subtitle: `${item.File_x0020_Type || 'Document'} • ${item.Author?.Title || 'Unknown'} • ${new Date(item.Modified).toLocaleDateString()}`,
                url: item.FileRef,
                fileType: item.File_x0020_Type || 'Document',
                lastModified: item.Modified,
                author: item.Author?.Title || 'Unknown'
              }));
              
              setSearchResults(results);
              setShowSuggestions(results.length > 0);
              setHighlightedIndex(-1);
              return;
            }
          }
        } catch (libraryError) {
          console.log(`Library ${libraryName} not accessible or empty:`, libraryError);
          continue;
        }
      }
      
      // No results found in any library
      setSearchResults([]);
      setShowSuggestions(false);
      
    } catch (error) {
      console.error('Other libraries search error:', error);
      throw error;
    }
  }, [context]);

  // Search document libraries directly
  const searchDocumentLibraries = useCallback(async (query: string) => {
    if (!context) return;
    
    try {
      // Search in the main Documents library
      const documentsUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Documents')/items?$filter=substringof('${encodeURIComponent(query)}',Title)&$select=Title,FileRef,FileLeafRef,Modified,Author/Title,File_x0020_Type&$expand=Author&$top=10&$orderby=Modified desc`;
      
      const response: SPHttpClientResponse = await context.spHttpClient.get(
        documentsUrl,
        SPHttpClient.configurations.v1
      );
      
      if (response.ok) {
        const data = await response.json();
        
        if (data.value && data.value.length > 0) {
          const results: ISharePointResult[] = data.value.map((item: any, index: number) => ({
            id: `doc-${index}`,
            title: item.Title || item.FileLeafRef || 'Untitled',
            subtitle: `${item.File_x0020_Type || 'Document'} • ${item.FileLeafRef?.split('.').pop() || 'Document'} • ${item.Author?.Title || 'Unknown'} • ${new Date(item.Modified).toLocaleDateString()}`,
            url: item.FileRef,
            fileType: item.File_x0020_Type || 'Document',
            lastModified: item.Modified,
            author: item.Author?.Title || 'Unknown'
          }));
          
          setSearchResults(results);
          setShowSuggestions(results.length > 0);
          setHighlightedIndex(-1);
          return;
        }
      }
      
      // If no results in Documents, try searching in other common libraries
      await searchOtherLibraries(query);
      
    } catch (error) {
      console.error('Documents library search error:', error);
      throw error;
    }
  }, [context, searchOtherLibraries]);

  // SharePoint search functionality - using simpler document library approach
  const searchSharePoint = useCallback(async (query: string) => {
    if (!query.trim() || !context) return;
    
    setIsSearching(true);
    try {
      // Use document library search which is more reliable than the search API
      await searchDocumentLibraries(query);
    } catch (error) {
      console.error('Document library search error:', error);
      setSearchResults([]);
      setShowSuggestions(false);
    } finally {
      setIsSearching(false);
    }
  }, [context, searchDocumentLibraries]);



  // Search on input change with debouncing
  useEffect(() => {
    if (!searchValue.trim() || !enableSuggestions) {
      setSearchResults([]);
      setShowSuggestions(false);
      return;
    }

    const timeoutId = setTimeout(() => {
      searchSharePoint(searchValue).catch(console.error);
    }, 300); // 300ms debounce

    return () => clearTimeout(timeoutId);
  }, [searchValue, enableSuggestions, searchSharePoint]);

  // Handle search execution
  const executeSearch = useCallback((searchQuery: string) => {
    if (searchQuery.trim()) {
      setShowSuggestions(false);
      setSearchValue(searchQuery);
      onSearch(searchQuery.trim());
    }
  }, [onSearch]);

  // Handle suggestion selection
  const selectSuggestion = useCallback((result: ISharePointResult) => {
    // Open the SharePoint document in a new tab
    if (result.url) {
      window.open(result.url, '_blank');
    }
    executeSearch(result.title);
  }, [executeSearch]);

  // Handle keyboard navigation
  const handleKeyDown = useCallback((event: React.KeyboardEvent) => {
    if (!showSuggestions || searchResults.length === 0) {
      if (event.key === 'Enter') {
        executeSearch(searchValue);
      }
      return;
    }

    switch (event.key) {
      case 'ArrowDown':
        event.preventDefault();
        setHighlightedIndex(prev => 
          prev < searchResults.length - 1 ? prev + 1 : 0
        );
        break;
      
      case 'ArrowUp':
        event.preventDefault();
        setHighlightedIndex(prev => 
          prev > 0 ? prev - 1 : searchResults.length - 1
        );
        break;
      
      case 'Enter':
        event.preventDefault();
        if (highlightedIndex >= 0 && highlightedIndex < searchResults.length) {
          selectSuggestion(searchResults[highlightedIndex]);
        } else {
          executeSearch(searchValue);
        }
        break;
      
      case 'Escape':
        setShowSuggestions(false);
        setHighlightedIndex(-1);
        break;
    }
  }, [showSuggestions, searchResults, highlightedIndex, searchValue, executeSearch, selectSuggestion]);

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
    if (searchValue.trim() && searchResults.length > 0) {
      setShowSuggestions(true);
    }
  }, [searchValue, searchResults]);

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
        autoComplete="off"
        aria-expanded={showSuggestions}
        aria-haspopup="listbox"
        role="combobox"
      />
      
      {showSuggestions && searchResults.length > 0 && (
        <div 
          className={styles.suggestionsDropdown}
          ref={suggestionsRef}
          role="listbox"
          aria-label="Search results"
        >
          {isSearching && (
            <div className={styles.searchingIndicator}>
              Searching...
            </div>
          )}
          {searchResults.map((result, index) => (
            <div
              key={result.id}
              className={`${styles.suggestionItem} ${index === highlightedIndex ? styles.highlighted : ''}`}
              onClick={() => selectSuggestion(result)}
              role="option"
              aria-selected={index === highlightedIndex}
              onMouseEnter={() => setHighlightedIndex(index)}
            >
              <div className={styles.suggestionText}>
                <div className={styles.suggestionTitle}>{result.title}</div>
                {result.subtitle && (
                  <div className={styles.suggestionSubtitle}>{result.subtitle}</div>
                )}
              </div>
            </div>
          ))}
        </div>
      )}
      
      {showSuggestions && searchResults.length === 0 && searchValue.trim() && !isSearching && (
        <div className={styles.suggestionsDropdown} ref={suggestionsRef}>
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
