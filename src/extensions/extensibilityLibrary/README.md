# My PnP Search Extensibility Library

A comprehensive SharePoint Framework (SPFx) extensibility library that integrates with PnP Modern Search web parts to provide enhanced search capabilities across SharePoint and Microsoft 365.

## 🚀 Features

- **Custom Data Sources**: SharePoint Search API with optional Microsoft Graph integration
- **Modern Layouts**: Responsive List and Card layout renderers  
- **Dynamic Filters**: Interactive refinement components with multi-select and operators
- **Rich Templates**: Handlebars templates with custom helpers and styling
- **Accessibility**: WCAG 2.1 compliant with keyboard navigation and screen reader support
- **Responsive Design**: Mobile-first approach with adaptive layouts
- **Performance**: Optimized with caching, lazy loading, and efficient API calls
- **TypeScript**: Fully typed interfaces and strict type checking
- **Error Handling**: Graceful error handling with user-friendly messages

## 📁 Project Structure

```
src/extensions/extensibilityLibrary/
├── components/
│   ├── dataSources/
│   │   ├── CustomSearchDataSource.ts
│   │   └── ICustomSearchDataSourceProps.ts
│   ├── layouts/
│   │   ├── ListLayout.tsx
│   │   ├── CardLayout.tsx
│   │   └── templates/
│   │       ├── ListLayout.html
│   │       └── CardLayout.html
│   └── filters/
│       └── CustomFilterComponent.tsx
├── services/
│   ├── SearchService.ts
│   └── GraphService.ts
├── common/
│   ├── interfaces/
│   │   └── index.ts
│   └── utils/
│       ├── SearchUtils.ts
│       └── GraphUtils.ts
├── styles/
│   └── common.scss
├── ExtensibilityLibrary.ts
├── IExtensibilityLibrary.ts
├── index.ts
└── README.md
```

## 🔧 Components

### Custom Data Sources

#### CustomSearchDataSource
- **SharePoint Search Integration**: Full KQL query support with refiners
- **Microsoft Graph Support**: Cross-platform search across OneDrive, Outlook, Teams
- **Configurable Properties**: Query templates, select properties, sorting, pagination
- **Caching**: Built-in result caching for improved performance
- **Error Handling**: Graceful fallbacks and error recovery

**Configuration Options:**
- Query Template (KQL)
- Select Properties
- Sort Fields
- Results Per Page
- Enable Graph Search
- Enable Refiners
- Cache Settings

### Custom Layouts

#### ListLayout
- Clean, accessible list view for search results
- Configurable metadata display
- Thumbnail support with file type icons
- Compact mode option
- Hover effects and animations
- Mobile-responsive design

**Configuration Options:**
- Show/Hide Metadata
- Show/Hide Thumbnails  
- Compact Mode
- Hover Effects
- Default Sort Order

#### CardLayout
- Modern card-based grid layout
- Responsive grid (1-6 cards per row)
- Rich metadata display
- Configurable card height and styling
- Shadow depth options
- Hover overlays with actions

**Configuration Options:**
- Cards Per Row (1-6)
- Card Height
- Border Radius
- Shadow Depth
- Show Author/Date
- Enable Animations

### Custom Filters

#### CustomFilterComponent
- Multi-select filter values
- AND/OR operator support
- Search within filters
- Collapsible panels
- Select All/Clear All actions
- Real-time filtering

**Features:**
- Dynamic filter loading
- Search within filter values
- Operator toggle (AND/OR)
- Accessibility keyboard navigation
- Mobile-responsive design

## 🛠️ Services

### SearchService
- **SharePoint Search API**: Complete REST API integration
- **KQL Query Building**: Advanced query construction utilities
- **Refinement Support**: Dynamic refiner management
- **Suggestion API**: Search query suggestions
- **Caching Layer**: Configurable result caching
- **Error Handling**: Comprehensive error management

### GraphService  
- **Microsoft Graph Integration**: Search across Microsoft 365
- **Multi-Entity Support**: Files, emails, events, contacts
- **Normalized Results**: Consistent result format
- **Permission Handling**: Graceful permission failures
- **Rate Limiting**: Built-in throttling support

## 🎨 Styling & Theming

### CSS Variables
- Fluent UI Design System compliance
- Dark theme support
- High contrast mode
- Responsive breakpoints
- Accessibility features

### Utility Classes
- Typography scales
- Color tokens
- Spacing system
- Flexbox utilities
- Responsive helpers

## 📚 Handlebars Helpers

Custom helpers provided for template rendering:

- `formatDate`: Format dates in user-friendly format
- `highlight`: Highlight search terms in results
- `truncate`: Truncate text with ellipsis
- `ifExists`: Conditional rendering for existing values

## 🚀 Getting Started

### Prerequisites
- SharePoint Online environment
- SPFx development environment
- Node.js 18+ 
- PnP Modern Search web parts installed

### Installation

1. **Build the Solution**
   ```bash
   npm install
   gulp build
   gulp bundle --ship
   gulp package-solution --ship
   ```

2. **Deploy to SharePoint**
   - Upload the `.sppkg` file to the SharePoint App Catalog
   - Install the app in your SharePoint site collection

3. **Use with PnP Modern Search**
   - Add a PnP Modern Search Results web part to a page
   - In the data source selection, choose "Custom SharePoint Search"
   - Configure layouts by selecting "Custom List Layout" or "Custom Card Layout"
   - Set up filters using the custom filter components

### Configuration

#### Data Source Configuration
```typescript
// Example configuration for CustomSearchDataSource
{
  queryTemplate: "ContentClass:STS_ListItem_DocumentLibrary",
  selectProperties: "Title,Path,Author,Created,Modified,Summary,FileType",
  sortBy: "Modified",
  rowLimit: 50,
  enableGraphSearch: true,
  enableRefiners: true,
  cacheDuration: 15
}
```

#### Layout Configuration
```typescript
// Example configuration for CardLayout
{
  cardsPerRow: 3,
  cardHeight: 300,
  showThumbnails: true,
  showAuthor: true,
  showDate: true,
  enableAnimations: true,
  shadowDepth: "medium",
  borderRadius: 8
}
```

## 🔌 API Reference

### ExtensibilityLibrary Class

Main entry point that implements `IExtensibilityLibrary`:

```typescript
export class ExtensibilityLibrary implements IExtensibilityLibrary {
  getCustomDataSources(): IDataSourceDefinition[]
  getCustomLayouts(): ILayoutDefinition[]
  getCustomWebComponents(): IComponentDefinition<any>[]
  getCustomSuggestionProviders(): ISuggestionProviderDefinition[]
  registerHandlebarsCustomizations(handlebarsNamespace: typeof Handlebars): void
  invokeCardAction(action: IAdaptiveCardAction): void
}
```

### Service Keys

```typescript
// Exported service keys for dependency injection
export const CustomSearchDataSourceServiceKey
export const ListLayoutServiceKey  
export const CardLayoutServiceKey
export const SearchServiceKey
export const GraphServiceKey
```

## 🧪 Development

### Building
```bash
npm run build
```

### Testing  
```bash
npm test
```

### Debugging
- Use `gulp serve` for local development
- Add `?debug=true` to test pages
- Check browser console for extensibility library logs

## 📝 Customization

### Extending Data Sources
Create new data sources by extending `BaseDataSource`:

```typescript
export class MyDataSource extends BaseDataSource<IMyDataSourceProps> {
  public async getData(dataContext?: IDataContext): Promise<IDataSourceData> {
    // Implementation
  }
}
```

### Custom Templates
Modify Handlebars templates in `components/layouts/templates/`:
- `ListLayout.html` - List view template
- `CardLayout.html` - Card view template

### Adding Utilities
Extend utility classes in `common/utils/`:
- `SearchUtils` - Search-related helpers
- `GraphUtils` - Microsoft Graph helpers

## 🐛 Troubleshooting

### Common Issues

1. **Library Not Loading**
   - Verify the .sppkg is deployed to App Catalog
   - Check if the app is installed in the site collection
   - Confirm PnP Modern Search web parts are installed

2. **Search Results Not Appearing**
   - Verify SharePoint Search service is working
   - Check query template syntax
   - Ensure proper permissions for Graph search

3. **Layouts Not Rendering**
   - Check browser console for JavaScript errors
   - Verify Handlebars template syntax
   - Ensure CSS files are loading properly

### Debug Mode
Enable debug logging by adding to browser console:
```javascript
localStorage.setItem('PnPModernSearchDebug', 'true');
```

## 📋 Requirements

- SharePoint Online
- SPFx 1.21.1+
- PnP Modern Search 4.x+
- Modern browsers (Chrome, Edge, Firefox, Safari)

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable  
5. Submit a pull request

## 📄 License

This project is licensed under the MIT License.

## 🆘 Support

For issues and questions:
- Check the troubleshooting section above
- Review PnP Modern Search documentation
- File issues in the project repository

## 🔗 References

- [PnP Modern Search](https://github.com/microsoft-search/pnp-modern-search)
- [PnP Modern Search Extensibility](https://microsoft-search.github.io/pnp-modern-search/extensibility/)
- [SharePoint Framework](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview)
- [Microsoft Graph](https://learn.microsoft.com/en-us/graph/)
- [Fluent UI](https://developer.microsoft.com/en-us/fluentui)

---

**Built with ❤️ for the SharePoint community**