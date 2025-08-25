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
import { useState, useCallback, useMemo, useEffect } from 'react';
import styles from './SpfxBannerSearch.module.scss';
import type { ISpfxBannerSearchProps } from './ISpfxBannerSearchProps';
// import { SearchBox } from '@fluentui/react/lib/SearchBox'; // Temporarily commented out for testing
import { ThemeProvider } from '@fluentui/react/lib/Theme';
import { Icon } from '@fluentui/react/lib/Icon'; // Still needed for AI toggle button
// Native SharePoint panels are now handled in the web part class

// Microsoft FluentUI official icon library (same as SharePoint/Office 365)
const ICON_BASE_URL = 'https://res.cdn.office.net/midgard/versionless/fluentui-resources/1.0.39/assets/item-types/24/';

const getFileTypeIconUrl = (extension: string): string => {
  const iconDict: { [key: string]: string } = {
    'eml': 'email.svg',
    'nws': 'genericfile.svg',
    'ascx': 'code.svg',
    'asp': 'code.svg',
    'aspx': 'spo.svg',
    'css': 'code.svg',
    'hta': 'genericfile.svg',
    'htm': 'html.svg',
    'html': 'html.svg',
    'htw': 'genericfile.svg',
    'htx': 'genericfile.svg',
    'jhtml': 'genericfile.svg',
    'stm': 'genericfile.svg',
    'mht': 'html.svg',
    'mhtml': 'html.svg',
    'xlb': 'sysfile.svg',
    'xlc': 'xlsx.svg',
    'xls': 'xlsx.svg',
    'xlsb': 'xlsx.svg',
    'xlsm': 'xlsx.svg',
    'xlsx': 'xlsx.svg',
    'xlt': 'xltx.svg',
    'one': 'one.svg',
    'pot': 'potx.svg',
    'ppa': 'sysfile.svg',
    'pps': 'ppsx.svg',
    'ppt': 'pptx.svg',
    'pptm': 'pptx.svg',
    'pptx': 'pptx.svg',
    'pub': 'pub.svg',
    'doc': 'docx.svg',
    'docm': 'docx.svg',
    'docx': 'docx.svg',
    'dot': 'dotx.svg',
    'dotx': 'dotx.svg',
    'xps': 'vector.svg',
    'odc': 'spreadsheet.svg',
    'odp': 'presentation.svg',
    'ods': 'spreadsheet.svg',
    'odt': 'rtf.svg',
    'msg': 'email.svg',
    'pdf': 'pdf.svg',
    'rtf': 'rtf.svg',
    'asm': 'code.svg',
    'bat': 'code.svg',
    'c': 'code.svg',
    'cmd': 'code.svg',
    'cpp': 'code.svg',
    'csv': 'csv.svg',
    'cxx': 'code.svg',
    'def': 'genericfile.svg',
    'h': 'code.svg',
    'hpp': 'code.svg',
    'lnk': 'link.svg',
    'mpx': 'genericfile.svg',
    'php': 'code.svg',
    'trf': 'genericfile.svg',
    'txt': 'txt.svg',
    'url': 'link.svg',
    'tif': 'photo.svg',
    'tiff': 'photo.svg',
    'jpg': 'photo.svg',
    'jpeg': 'photo.svg',
    'png': 'photo.svg',
    'gif': 'photo.svg',
    'bmp': 'photo.svg',
    'vdw': 'vsdx.svg',
    'vdx': 'vsdx.svg',
    'vsd': 'vsdx.svg',
    'vsdm': 'vsdx.svg',
    'vsdx': 'vsdx.svg',
    'vss': 'vssx.svg',
    'vssm': 'vssx.svg',
    'vssx': 'vssx.svg',
    'vst': 'vstx.svg',
    'vstm': 'vstx.svg',
    'vsx': 'vstx.svg',
    'vtx': 'genericfile.svg',
    'jsp': 'code.svg',
    'mspx': 'genericfile.svg',
    'rss': 'code.svg',
    'xml': 'xml.svg',
    'zip': 'zip.svg',
    'rar': 'zip.svg',
    '7z': 'zip.svg',
    'mp4': 'video.svg',
    'avi': 'video.svg',
    'mov': 'video.svg',
    'wmv': 'video.svg',
    'mp3': 'audio.svg',
    'wav': 'audio.svg',
    'wma': 'audio.svg'
  };

  if (extension && extension.trim() !== '') {
    const lowerExt = extension.toLowerCase();
    const iconFile = iconDict[lowerExt] || 'genericfile.svg';
    return `${ICON_BASE_URL}${iconFile}`;
  } else {
    return `${ICON_BASE_URL}docset.svg`;
  }
};
import AISearch from './AISearch';
// ...existing code...
import { WebPartContext } from '@microsoft/sp-webpart-base';

import { SharePointSearchService, SuggestionItem } from '../../../services/SharePointSearchService';
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
  suggestionsLimit: number;
  openingBehavior: string;
  enableQuerySuggestions: boolean;
  staticSuggestions: string;
  enableZeroTermSuggestions: boolean;
  zeroTermSuggestions: string;
  suggestionsProvider: string;
  hubSiteId: string;
  imageRelativeUrl: string;
  searchBoxBorderRadius: number;
  searchBoxHeight: number;
  useDynamicDataSource: boolean;
  dynamicDataSourceId: string;
  pageEnvironmentProperty: string;
  siteProperty: string;
  userProperty: string;
  queryStringProperty: string;
  searchProperty: string;
  semanticColors: Partial<import('@fluentui/react/lib/Styling').ISemanticColors>;
  context: WebPartContext;
  searchSiteUrl?: string;
  debugSuggestions?: boolean;
}> = React.memo(({ placeholder, onSearch, enableSuggestions, suggestionsLimit, openingBehavior, enableQuerySuggestions, staticSuggestions, enableZeroTermSuggestions, zeroTermSuggestions, suggestionsProvider, hubSiteId, imageRelativeUrl, searchBoxBorderRadius, searchBoxHeight, useDynamicDataSource, dynamicDataSourceId, pageEnvironmentProperty, siteProperty, userProperty, queryStringProperty, searchProperty, semanticColors, context, searchSiteUrl, debugSuggestions }) => {
  console.debug("[HeroSearchBox] Component is rendering with props:", { placeholder, enableSuggestions });
  const service = useMemo(() => new SharePointSearchService(context, searchSiteUrl, debugSuggestions), [context, searchSiteUrl, debugSuggestions]);

  // Dynamic data source function
  const getDynamicDataValue = useCallback((): string => {
    if (!useDynamicDataSource || dynamicDataSourceId !== 'pageEnvironment') {
      return '';
    }

    switch (pageEnvironmentProperty) {
      case 'siteProperties':
        switch (siteProperty) {
          case 'siteUrl':
            return context.pageContext.web.absoluteUrl;
          case 'siteCollectionUrl':
            return context.pageContext.site.absoluteUrl;
          case 'siteTitle':
            return context.pageContext.web.title;
          case 'siteId':
            return context.pageContext.site.id.toString();
          case 'webId':
            return context.pageContext.web.id.toString();
          case 'hubSiteId':
            // Note: hubSiteId is not available in basic pageContext, would need Graph API call
            return '';
          default:
            return '';
        }
      case 'currentUser':
        switch (userProperty) {
          case 'loginName':
            return context.pageContext.user.loginName;
          case 'displayName':
            return context.pageContext.user.displayName;
          case 'email':
            return context.pageContext.user.email;
          case 'userId':
            return context.pageContext.user.loginName; // Use loginName as userId
          case 'department':
            return ''; // Not available in basic page context
          case 'jobTitle':
            return ''; // Not available in basic page context
          default:
            return '';
        }
      case 'queryString':
        if (queryStringProperty) {
          const urlParams = new URLSearchParams(window.location.search);
          return urlParams.get(queryStringProperty) || '';
        }
        return '';
      case 'search':
        // This would typically come from connected search web parts
        // For now, return empty as it requires more complex implementation
        return '';
      default:
        return '';
    }
  }, [useDynamicDataSource, dynamicDataSourceId, pageEnvironmentProperty, siteProperty, userProperty, queryStringProperty, searchProperty, context]);
  
  // Create zero-term suggestions
  const zeroTermSuggestionsItems = useMemo(() => {
    if (!enableQuerySuggestions || !enableZeroTermSuggestions) return [];
    return service.getZeroTermSuggestions(zeroTermSuggestions);
  }, [service, enableQuerySuggestions, enableZeroTermSuggestions, zeroTermSuggestions]);

  // Create a stable fetchFn to prevent infinite loops
  const fetchFn = useCallback(
    async (term: string, signal?: AbortSignal, limit?: number): Promise<SuggestionItem[]> => {
      if (!enableSuggestions && !enableQuerySuggestions) return [];
      
      const allSuggestions: SuggestionItem[] = [];
      
      // Add file suggestions (existing behavior)
      if (enableSuggestions && term.trim()) {
        try {
          // Pass hub site ID if using custom provider
          const hubId = (enableQuerySuggestions && suggestionsProvider === 'custom') ? hubSiteId : undefined;
          const fileSuggestions = await service.fetchSuggestions(term, signal, limit, hubId);
          allSuggestions.push(...fileSuggestions);
        } catch (error) {
          console.warn("[HeroSearchBox] File suggestions failed:", error);
        }
      }
      
      // Add query suggestions (new feature)
      if (enableQuerySuggestions && suggestionsProvider === 'static') {
        const querySuggestions = service.getStaticSuggestions(staticSuggestions, term);
        allSuggestions.push(...querySuggestions);
      }
      
      // Limit total suggestions
      return allSuggestions.slice(0, limit || suggestionsLimit);
    },
    [service, enableSuggestions, enableQuerySuggestions, suggestionsProvider, staticSuggestions, suggestionsLimit, hubSiteId]
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
  } = useTypeahead(fetchFn, 250, suggestionsLimit, zeroTermSuggestionsItems);

  // Auto-populate search box with dynamic data
  useEffect(() => {
    if (useDynamicDataSource && pageEnvironmentProperty && (siteProperty || userProperty || queryStringProperty || searchProperty)) {
      const dynamicValue = getDynamicDataValue();
      if (dynamicValue && dynamicValue !== searchValue) {
        console.debug("[HeroSearchBox] Auto-populating with dynamic data:", dynamicValue);
        onChange(dynamicValue);
      }
    }
  }, [useDynamicDataSource, pageEnvironmentProperty, siteProperty, userProperty, queryStringProperty, searchProperty, getDynamicDataValue, onChange, searchValue]);
  
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

  // State for input focus to control magnifying glass icon
  const [isInputFocused, setIsInputFocused] = useState<boolean>(false);

  // Handle input focus - hide magnifying glass
  const handleFocus = useCallback((): void => {
    setIsInputFocused(true);
    // Focus handling is now managed by useTypeahead hook
  }, []);

  // Handle input blur - show magnifying glass again
  const handleInputBlur = useCallback((event: React.FocusEvent) => {
    setIsInputFocused(false);
    // Delay hiding suggestions to allow for suggestion clicks
    setTimeout(() => {
      if (event.currentTarget && !event.currentTarget.contains(document.activeElement)) {
        setShowSuggestions(false);
        setHighlightedIndex(-1);
      }
    }, 150);
  }, [setShowSuggestions]);

  // Note: handleBlur replaced by handleInputBlur above for better focus management

  console.debug("[HeroSearchBox] Rendering SearchBox with value:", searchValue);
  
  return (
    <div className={styles.searchContainer}>
      {/* Search input with magnifying glass icon */}
      <div style={{ 
        position: 'relative', 
        width: '100%',
        display: 'flex',
        alignItems: 'center'
      }}>
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
          onBlur={handleInputBlur}
          className={styles.heroSearchBox}
          autoComplete="off"
          aria-expanded={showSuggestions}
          aria-haspopup="listbox"
          role="combobox"
          style={{
            width: '100%',
            border: 'none',
            fontSize: '1rem',
            padding: '0px 10px 1.5px',
            paddingLeft: isInputFocused ? '10px' : '35px',
            borderRadius: `${searchBoxBorderRadius}px`,
            outline: 'none',
            height: `${searchBoxHeight}px`,
            transition: 'padding-left 0.3s ease-in-out'
          }}
        />
        <Icon
          iconName="Search"
          style={{
            position: 'absolute',
            left: '12px',
            fontSize: '16px',
            color: '#666',
            pointerEvents: 'none',
            opacity: isInputFocused ? 0 : 1,
            transform: isInputFocused ? 'scale(0.8)' : 'scale(1)',
            transition: 'opacity 0.3s ease-in-out, transform 0.3s ease-in-out'
          }}
        />
      </div>
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
                // If item has a path, open it based on configured behavior, otherwise perform search
                if (item.path && item.path.trim()) {
                  if (openingBehavior === 'current-tab') {
                    window.location.href = item.path;
                  } else {
                    // Default to new tab ('new-tab')
                    window.open(item.path, '_blank', 'noopener,noreferrer');
                  }
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
              <img
                src={getFileTypeIconUrl(item.fileType || '')}
                alt={`${item.fileType || 'file'} icon`}
                className={styles.fileIcon}
                style={{ 
                  width: '16px', 
                  height: '16px', 
                  marginRight: '8px',
                  flexShrink: 0
                }}
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
    bannerTitleColor,
    bannerTitle,
    searchBoxPlaceholder,
    enableSuggestions,
    isDarkTheme,
    semanticColors,
    onSearchQuery,
    hasTeamsContext,
    context
  } = props;

  // Parse extensibility libraries for demo effects
  // Extensibility libraries parsing removed - indicator not needed

  // Local state for AI search toggle (defaults to false, users can toggle)
  const [isAISearchActive, setIsAISearchActive] = useState<boolean>(false);
  
  // Panel states removed - now using native SharePoint panels

  // Memoized styles for performance
  const bannerStyles = useMemo(() => ({
    '--gradient-start': gradientStartColor,
    '--gradient-end': gradientEndColor,
    '--min-height': `${minHeight}px`,
    '--title-font-size': `${titleFontSize}px`,
    '--title-color': bannerTitleColor,
    '--body-text': semanticColors?.bodyText || (isDarkTheme ? '#ffffff' : '#323130'),
    '--link-color': semanticColors?.link || '#0078d4',
    '--link-hover': semanticColors?.linkHovered || '#106ebe'
  } as React.CSSProperties), [gradientStartColor, gradientEndColor, minHeight, titleFontSize, bannerTitleColor, semanticColors, isDarkTheme]);

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

  // Panel handling now done in web part class with native SharePoint panels

  // Panel logic removed - now using native SharePoint panels in web part class

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
        
        {/* Extensibility Libraries Indicator - removed */}
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
                suggestionsLimit={props.suggestionsLimit}
                openingBehavior={props.openingBehavior}
                enableQuerySuggestions={props.enableQuerySuggestions}
                staticSuggestions={props.staticSuggestions}
                enableZeroTermSuggestions={props.enableZeroTermSuggestions}
                zeroTermSuggestions={props.zeroTermSuggestions}
                suggestionsProvider={props.suggestionsProvider}
                hubSiteId={props.hubSiteId}
                imageRelativeUrl={props.imageRelativeUrl}
                searchBoxBorderRadius={props.searchBoxBorderRadius}
                searchBoxHeight={props.searchBoxHeight}
                useDynamicDataSource={props.useDynamicDataSource}
                dynamicDataSourceId={props.dynamicDataSourceId}
                pageEnvironmentProperty={props.pageEnvironmentProperty}
                siteProperty={props.siteProperty}
                userProperty={props.userProperty}
                queryStringProperty={props.queryStringProperty}
                searchProperty={props.searchProperty}
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

      {/* Panels are now handled as native SharePoint panels in the web part class */}
    </ThemeProvider>
  );
};

console.debug("[SpfxBannerSearch] Module loaded");

export default SpfxBannerSearch;
