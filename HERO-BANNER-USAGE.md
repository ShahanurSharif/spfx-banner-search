# SPFx Hero Banner Search Web Part - Usage Guide

## Overview

The **AI Search Hero Banner** is a full-bleed SPFx web part that provides a visually stunning search interface for SharePoint pages. It features configurable gradients, optional animations, and seamless integration with PnP Modern Search through dynamic data connections.

## Key Features

### 🎨 Visual Design
- **Full-bleed support**: Expands to full page width when placed in a full-width column
- **Configurable gradients**: Customizable start and end colors for the background
- **Animated circles**: Optional floating circle animations for visual appeal
- **Responsive design**: Adapts to different screen sizes and devices
- **Theme-aware**: Automatically adjusts to SharePoint themes and dark mode

### 🔍 Search Functionality
- **Centralized search box**: Large, prominent search input with modern styling
- **Dynamic data publishing**: Publishes search queries for consumption by other web parts
- **Suggestions support**: Ready for integration with search suggestion providers
- **PnP Modern Search integration**: Compatible with PnP Modern Search Results web parts

### ⚡ Performance & Accessibility
- **Optimized rendering**: Uses React.memo, useMemo, and useCallback for performance
- **Accessibility compliant**: Proper ARIA labels, keyboard navigation, and screen reader support
- **High contrast support**: Adapts to Windows high contrast mode
- **Print-friendly**: Optimized styles for printing

## Configuration

### Step 1: Banner & Visual Settings

#### Gradient Colors
- **Gradient Start Color**: The starting color of the background gradient (default: #0078d4)
- **Gradient End Color**: The ending color of the background gradient (default: #106ebe)

#### Visual Elements
- **Show Circle Animation**: Enable/disable floating circle animations behind the search box
- **Banner Minimum Height**: Height of the banner in pixels (300-800px, default: 450px)
- **Search Box Placeholder**: Text displayed in the search box when empty

### Step 2: Search Configuration

#### Query Settings
- **Query Template**: Default query template to use (e.g., "*" for all results)
- **Results Web Part ID**: Optional GUID of a Search Results web part to connect to
- **Enable Search Suggestions**: Toggle search suggestions functionality

## Integration with PnP Modern Search

### Dynamic Data Connection

The web part publishes a dynamic data property called `inputQueryText` that can be consumed by PnP Modern Search Results web parts:

1. **Add the Hero Banner web part** to your page
2. **Add a PnP Modern Search Results web part** below it
3. **Configure the Search Results web part**:
   - Go to the "Connections" configuration page
   - Select "Connect to a dynamic data source"
   - Choose the Hero Banner web part as the source
   - Select "Search Query Text" as the property
4. **Set up the query**: Use the `{inputQueryText}` token in your search query template

### Example Connection Setup

```
Search Results Web Part Configuration:
├── Data Source: SharePoint Search
├── Query Template: {inputQueryText}
├── Input Query Text: 
│   ├── Source: Hero Banner Web Part
│   └── Property: Search Query Text
└── Default Query: * (for initial results)
```

## Layout Requirements

### Full-Bleed Setup

To enable full-bleed mode:

1. **Edit the page** where you want to add the web part
2. **Add a section** and choose "Full-width column"
3. **Add the AI Search Hero Banner web part** to this section
4. The web part will automatically expand to full page width

### Standard Layout

The web part also works in standard column layouts but will be constrained to the column width.

## Styling & Theming

### CSS Variables

The component uses CSS variables for dynamic theming:

```scss
--gradient-start: /* Start color of gradient */
--gradient-end: /* End color of gradient */
--min-height: /* Minimum height in pixels */
--body-text: /* Text color from theme */
--link-color: /* Link color from theme */
--link-hover: /* Link hover color from theme */
```

### Theme Integration

The web part automatically:
- Consumes SharePoint theme colors through the ThemeProvider
- Adapts to light/dark mode preferences
- Supports high contrast mode
- Maintains brand consistency across different tenants

## Browser Support

- ✅ Chrome (latest)
- ✅ Edge Chromium (latest)
- ✅ Firefox (latest)
- ✅ Safari (latest)
- ✅ Mobile browsers (responsive design)

## Performance Considerations

### Optimization Features
- **React.memo**: Prevents unnecessary re-renders of child components
- **useMemo**: Caches expensive style calculations
- **useCallback**: Optimizes event handlers
- **CSS Hardware Acceleration**: Smooth animations using transform and opacity
- **Lazy Loading**: Animation elements only render when enabled

### Best Practices
- Use the web part once per page for optimal performance
- Test with various gradient colors for accessibility compliance
- Consider disabling animations on slower devices if needed

## Accessibility Features

### ARIA Support
- Proper semantic HTML structure
- Search landmarks for screen readers
- Descriptive labels and roles
- Keyboard navigation support

### Visual Accessibility
- High contrast mode support
- Color-blind friendly gradient options
- Scalable text and elements
- Focus indicators for keyboard users

## Troubleshooting

### Common Issues

**Q: The web part doesn't span full width**
A: Ensure you've placed it in a "Full-width column" section and that `supportsFullBleed: true` is set in the manifest.

**Q: Search queries aren't reaching the Results web part**
A: Check the dynamic data connection configuration and ensure the Results web part is configured to consume the `inputQueryText` property.

**Q: Animations are causing performance issues**
A: Disable the "Show Circle Animation" option in the web part configuration.

**Q: Colors don't match my theme**
A: The web part uses configurable gradient colors. Adjust them to match your brand or let the automatic theme integration handle it.

### Debug Mode

To enable debug logging:
1. Open browser developer tools
2. Set localStorage item: `localStorage.setItem('spfx-debug', 'true')`
3. Reload the page

## Future Enhancements

Planned features for future versions:
- Advanced suggestion providers integration
- More animation options
- Video background support
- Custom search scopes
- Analytics integration
- Multi-language support

## Support

For issues, questions, or feature requests:
1. Check the troubleshooting section above
2. Review the PnP Modern Search documentation
3. Open an issue in the project repository
4. Contact the development team

---

*Built with SharePoint Framework v1.21.1 and PnP Modern Search Extensibility*
