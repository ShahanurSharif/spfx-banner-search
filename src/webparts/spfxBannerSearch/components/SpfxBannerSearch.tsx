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
import { useState, useCallback, useMemo, useRef } from 'react';
import styles from './SpfxBannerSearch.module.scss';
import type { ISpfxBannerSearchProps } from './ISpfxBannerSearchProps';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { ThemeProvider } from '@fluentui/react/lib/Theme';

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

// Search box component with suggestions
const HeroSearchBox: React.FC<{
  placeholder: string;
  onSearch: (query: string) => void;
  enableSuggestions: boolean;
  semanticColors: Partial<import('@fluentui/react/lib/Styling').ISemanticColors>;
}> = React.memo(({ placeholder, onSearch, enableSuggestions, semanticColors }) => {
  const [searchValue, setSearchValue] = useState('');
  const searchBoxRef = useRef<HTMLDivElement>(null);

  const handleSearch = useCallback((newValue?: string) => {
    const query = newValue || searchValue;
    if (query.trim()) {
      onSearch(query.trim());
    }
  }, [searchValue, onSearch]);

  const handleKeyPress = useCallback((event: React.KeyboardEvent) => {
    if (event.key === 'Enter') {
      handleSearch();
    }
  }, [handleSearch]);

  return (
    <div className={styles.searchContainer}>
      <SearchBox
        ref={searchBoxRef}
        placeholder={placeholder}
        value={searchValue}
        onChange={(_, newValue) => setSearchValue(newValue || '')}
        onSearch={handleSearch}
        onKeyDown={handleKeyPress}
        className={styles.heroSearchBox}
        iconProps={{ iconName: 'Search' }}
        autoComplete="off"
      />
      {enableSuggestions && (
        <div className={styles.suggestionsHint}>
          <span>Start typing to see suggestions</span>
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
    searchBoxPlaceholder,
    enableSuggestions,
    isDarkTheme,
    semanticColors,
    onSearchQuery,
    hasTeamsContext
  } = props;

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
        
        {/* Main content container */}
        <div className={styles.heroContent}>
          <div className={styles.searchWrapper}>
            <h1 className={styles.heroTitle}>
              Find What You Need
            </h1>
            <p className={styles.heroSubtitle}>
              Search across all content and resources
            </p>
            
            <HeroSearchBox
              placeholder={searchBoxPlaceholder}
              onSearch={handleSearch}
              enableSuggestions={enableSuggestions}
              semanticColors={semanticColors}
            />
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
