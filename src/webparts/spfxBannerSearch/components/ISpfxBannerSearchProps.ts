import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ISemanticColors } from '@fluentui/react/lib/Styling';

export interface ISpfxBannerSearchProps {
  // Banner & Search Box configuration
  gradientStartColor: string;
  gradientEndColor: string;
  showCircleAnimation: boolean;
  minHeight: number;
  titleFontSize: number;
  bannerTitle: string;
  searchBoxPlaceholder: string;
  
  // Search behavior configuration
  queryTemplate: string;
  enableSuggestions: boolean;
  suggestionsLimit: number;
  openingBehavior: string;
  
  // SPFx context and theme
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  semanticColors: Partial<ISemanticColors>;
  context: WebPartContext;
  onSearchQuery: (queryText: string) => void;

  /**
   * Optional: Override the site URL for search suggestions (defaults to context.pageContext.web.absoluteUrl)
   */
  searchSiteUrl?: string;

  /**
   * Optional: Enable debug logging for type-ahead suggestions
   */
  debugSuggestions?: boolean;
}
