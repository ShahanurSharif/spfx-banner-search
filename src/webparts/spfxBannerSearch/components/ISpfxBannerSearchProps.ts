import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ISemanticColors } from '@fluentui/react/lib/Styling';

export interface ISpfxBannerSearchProps {
  // Banner & Search Box configuration
  gradientStartColor: string;
  gradientEndColor: string;
  showCircleAnimation: boolean;
  minHeight: number;
  searchBoxPlaceholder: string;
  
  // Search behavior configuration
  queryTemplate: string;
  enableSuggestions: boolean;
  enableAISearch: boolean;
  
  // SPFx context and theme
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  semanticColors: Partial<ISemanticColors>;
  context: WebPartContext;
  onSearchQuery: (queryText: string) => void;
}
