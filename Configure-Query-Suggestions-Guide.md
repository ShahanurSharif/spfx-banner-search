# Configure Query Suggestions Guide

## Overview
The Query Suggestions feature provides intelligent search recommendations that combine file search results with configurable query suggestions, similar to PnP Modern Search extensibility. This guide explains how to configure and use all the query suggestion options.

## Configuration Location
Navigate to **Web Part Properties** → **Query Suggestions** (Third Configuration Page)

## Configuration Options

### 1. Enable Query Suggestions
**Toggle:** `Enable Query Suggestions`
- **Purpose:** Master switch for all query suggestion functionality
- **Default:** ON
- **Behavior:** 
  - ✅ **ON:** Shows both file results and query suggestions
  - ❌ **OFF:** No suggestion box appears at all (overrides all other settings)

### 2. Suggestions Provider
**Dropdown:** `Suggestions Provider`
- **Options:**
  - `SharePoint Static Suggestions` - Use predefined suggestions
  - `SharePoint Search Suggestions` - Use SharePoint's search API (future feature)
  - `Custom Provider` - External suggestion source (future feature)
- **Default:** `SharePoint Static Suggestions`
- **Current Implementation:** Only static suggestions are fully implemented

### 3. Static Suggestions
**Multi-line Text Field:** `Static Suggestions`
- **Purpose:** Define custom search suggestions that appear when users type
- **Format:** Comma-separated values
- **Example:** `policy, procedure, guidelines, training, safety, HR portal`
- **Behavior:** 
  - Filters suggestions based on user input
  - Shows maximum 10 suggestions
  - Case-insensitive matching

### 4. Zero-term Suggestions
**Toggle:** `Show suggestions when search box is empty`
- **Purpose:** Display popular searches when the search box is clicked but empty
- **Default:** ON
- **Dependencies:** Requires `Enable Query Suggestions` to be ON

**Multi-line Text Field:** `Zero-term Suggestions`
- **Purpose:** Define suggestions shown in empty search box
- **Format:** Comma-separated values
- **Example:** `help desk, IT support, HR portal, company policies, training materials`
- **Behavior:** Shows maximum 5 suggestions when search box is empty

## User Experience Flows

### Typing Behavior
1. **User types "pol"**
   - Shows: `policy.pdf` (file result) + `policy` (query suggestion)
   - File results come from SharePoint search
   - Query suggestions filter from static list

2. **User types "help"**
   - Shows: `help.docx` (file result) + `help desk` (query suggestion)
   - Combined results for comprehensive search experience

### Empty Search Box Behavior
1. **User clicks empty search box**
   - Shows: `help desk, IT support, HR portal` (zero-term suggestions)
   - Provides quick access to popular searches
   - Only appears if zero-term suggestions are configured

### Click Behaviors
1. **Clicking File Results**
   - Opens the actual file/document
   - Respects "Opening behavior" setting (current tab vs new tab)

2. **Clicking Query Suggestions**
   - Performs a search for that term
   - Fills search box and triggers search action
   - Follows redirect settings if configured

## Configuration Examples

### Basic Setup
```
✅ Enable Query Suggestions: ON
📋 Suggestions Provider: SharePoint Static Suggestions
📝 Static Suggestions: policy, procedure, training, safety
✅ Show suggestions when search box is empty: ON
📝 Zero-term Suggestions: help, support, training
```

### Corporate Portal Setup
```
✅ Enable Query Suggestions: ON
📋 Suggestions Provider: SharePoint Static Suggestions
📝 Static Suggestions: employee handbook, benefits, payroll, IT support, facilities, training, compliance, safety procedures, org chart, contact directory
✅ Show suggestions when search box is empty: ON
📝 Zero-term Suggestions: help desk, HR portal, IT support, employee benefits, training catalog
```

### Department-Specific Setup
```
✅ Enable Query Suggestions: ON
📋 Suggestions Provider: SharePoint Static Suggestions
📝 Static Suggestions: project templates, design guidelines, brand assets, marketing materials, campaign data, customer insights
✅ Show suggestions when search box is empty: ON
📝 Zero-term Suggestions: brand guidelines, templates, reports
```

## Toggle States Reference

| Enable Query Suggestions | Enable File Suggestions | Result |
|-------------------------|-------------------------|---------|
| ❌ OFF | ❌ OFF | **No suggestions at all** |
| ❌ OFF | ✅ ON | **Only file suggestions** |
| ✅ ON | ❌ OFF | **Only query suggestions** |
| ✅ ON | ✅ ON | **Both types combined** ✅ |

## Best Practices

### Suggestion Content
- **Use clear, searchable terms** that users actually look for
- **Include department-specific keywords** relevant to your audience
- **Keep suggestions concise** - avoid long phrases
- **Test with real users** to validate suggestion relevance

### Static Suggestions Format
```
✅ Good: policy, procedure, training, safety, benefits
❌ Bad: policy documents, procedure manuals, training materials and courses
```

### Zero-term Suggestions
- **Limit to 3-5 items** for better user experience
- **Focus on most popular searches** in your organization
- **Use action-oriented terms** like "help desk" rather than just "help"

### Performance Considerations
- **Static suggestions are instant** - no API calls required
- **File suggestions may have latency** - depends on SharePoint search performance
- **Combined results are optimized** - suggestions load in parallel

## Troubleshooting

### No Suggestions Appearing
1. Check if `Enable Query Suggestions` is ON
2. Verify `Enable File Suggestions` setting (first page)
3. Ensure Static Suggestions field has content
4. Check browser console for errors

### Suggestions Not Filtering
1. Verify comma-separated format in Static Suggestions
2. Check for extra spaces or special characters
3. Test with simple terms first

### Zero-term Suggestions Not Showing
1. Ensure `Show suggestions when search box is empty` is ON
2. Verify Zero-term Suggestions field has content
3. Check that `Enable Query Suggestions` is ON

## Future Enhancements
- **SharePoint Search Suggestions:** Dynamic suggestions from SharePoint search API
- **Custom Provider Support:** Integration with external suggestion services
- **Analytics Integration:** Track popular searches for suggestion optimization
- **Advanced Filtering:** More sophisticated suggestion matching algorithms

## Technical Notes
- Suggestions are case-insensitive
- Maximum 10 query suggestions + unlimited file results
- Zero-term suggestions limited to 5 items
- All suggestions respect the overall suggestion limit setting
- File results always take priority in display order

---
*This feature implements PnP Modern Search style query suggestions with enhanced configurability and user experience optimization.*
