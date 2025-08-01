/**
 * Custom Filter Component
 * 
 * Provides dynamic refinement filtering for search results
 * Supports multiple filter types, operators, and interactive UI
 */

import * as React from 'react';
import { useState, useCallback, useEffect } from 'react';
import { ExtensibilityConstants } from '@pnp/modern-search-extensibility';

/**
 * Filter value interface
 */
export interface IFilterValue {
    name: string;
    value: string;
    count: number;
    selected: boolean;
}

/**
 * Filter configuration interface
 */
export interface IFilter {
    filterName: string;
    displayName: string;
    values: IFilterValue[];
    operator: 'OR' | 'AND';
    maxValues?: number;
    showSearch?: boolean;
    collapsed?: boolean;
    showCount?: boolean;
}

/**
 * Custom Filter Component Properties
 */
export interface ICustomFilterComponentProps {
    filters: IFilter[];
    onFilterUpdate: (filterName: string, selectedValues: string[]) => void;
    onOperatorUpdate: (filterName: string, operator: 'OR' | 'AND') => void;
    onApplyAll: () => void;
    onClearAll: () => void;
    theme?: 'light' | 'dark';
    compact?: boolean;
    showOperatorToggle?: boolean;
}

/**
 * Individual Filter Component
 */
const FilterPanel: React.FC<{
    filter: IFilter;
    onFilterUpdate: (selectedValues: string[]) => void;
    onOperatorUpdate: (operator: 'OR' | 'AND') => void;
    theme?: 'light' | 'dark';
    compact?: boolean;
    showOperatorToggle?: boolean;
}> = ({ filter, onFilterUpdate, onOperatorUpdate, theme = 'light', compact = false, showOperatorToggle = true }) => {
    
    const [collapsed, setCollapsed] = useState(filter.collapsed || false);
    const [searchTerm, setSearchTerm] = useState('');
    const [filteredValues, setFilteredValues] = useState<IFilterValue[]>(filter.values);

    // Filter values based on search term
    useEffect(() => {
        if (searchTerm) {
            setFilteredValues(
                filter.values.filter(value =>
                    value.name.toLowerCase().includes(searchTerm.toLowerCase())
                )
            );
        } else {
            setFilteredValues(filter.values);
        }
    }, [searchTerm, filter.values]);

    const handleValueToggle = useCallback((value: string, selected: boolean) => {
        const currentSelected = filter.values
            .filter(v => v.selected)
            .map(v => v.value);

        let newSelected: string[];
        if (selected) {
            newSelected = [...currentSelected, value];
        } else {
            newSelected = currentSelected.filter(v => v !== value);
        }

        // Apply max values limit
        if (filter.maxValues && newSelected.length > filter.maxValues) {
            newSelected = newSelected.slice(-filter.maxValues);
        }

        onFilterUpdate(newSelected);
    }, [filter, onFilterUpdate]);

    const handleSelectAll = useCallback(() => {
        const allValues = filteredValues.map(v => v.value);
        const limitedValues = filter.maxValues ? allValues.slice(0, filter.maxValues) : allValues;
        onFilterUpdate(limitedValues);
    }, [filteredValues, filter.maxValues, onFilterUpdate]);

    const handleClearAll = useCallback(() => {
        onFilterUpdate([]);
    }, [onFilterUpdate]);

    const selectedCount = filter.values.filter(v => v.selected).length;
    const hasSelection = selectedCount > 0;

    return (
        <div className={`filter-panel ${theme} ${compact ? 'compact' : ''}`}>
            
            {/* Filter Header */}
            <div className="filter-header" onClick={() => setCollapsed(!collapsed)}>
                <h4 className="filter-title">
                    {filter.displayName}
                    {hasSelection && (
                        <span className="selection-count">({selectedCount})</span>
                    )}
                </h4>
                <button 
                    type="button" 
                    className="collapse-button"
                    aria-label={collapsed ? 'Expand filter' : 'Collapse filter'}
                >
                    <i className={`ms-Icon ms-Icon--${collapsed ? 'ChevronDown' : 'ChevronUp'}`} aria-hidden="true" />
                </button>
            </div>

            {/* Filter Content */}
            {!collapsed && (
                <div className="filter-content">
                    
                    {/* Operator Toggle */}
                    {showOperatorToggle && hasSelection && (
                        <div className="filter-operator">
                            <span className="operator-label">Match:</span>
                            <div className="operator-toggle">
                                <button
                                    type="button"
                                    className={`operator-button ${filter.operator === 'OR' ? 'active' : ''}`}
                                    onClick={() => onOperatorUpdate('OR')}
                                    title="Match any selected values"
                                >
                                    Any
                                </button>
                                <button
                                    type="button"
                                    className={`operator-button ${filter.operator === 'AND' ? 'active' : ''}`}
                                    onClick={() => onOperatorUpdate('AND')}
                                    title="Match all selected values"
                                >
                                    All
                                </button>
                            </div>
                        </div>
                    )}

                    {/* Search Box */}
                    {filter.showSearch && filter.values.length > 5 && (
                        <div className="filter-search">
                            <input
                                type="text"
                                placeholder="Search filters..."
                                value={searchTerm}
                                onChange={(e) => setSearchTerm(e.target.value)}
                                className="search-input"
                            />
                            <i className="ms-Icon ms-Icon--Search search-icon" aria-hidden="true" />
                        </div>
                    )}

                    {/* Action Buttons */}
                    {filter.values.length > 3 && (
                        <div className="filter-actions">
                            <button
                                type="button"
                                className="action-button select-all"
                                onClick={handleSelectAll}
                                disabled={filteredValues.length === 0}
                            >
                                Select All
                            </button>
                            <button
                                type="button"
                                className="action-button clear-all"
                                onClick={handleClearAll}
                                disabled={!hasSelection}
                            >
                                Clear All
                            </button>
                        </div>
                    )}

                    {/* Filter Values */}
                    <div className="filter-values">
                        {filteredValues.length > 0 ? (
                            filteredValues.map((value) => (
                                <label key={value.value} className="filter-value-item">
                                    <input
                                        type="checkbox"
                                        checked={value.selected}
                                        onChange={(e) => handleValueToggle(value.value, e.target.checked)}
                                        className="value-checkbox"
                                    />
                                    <span className="value-label">{value.name}</span>
                                    {filter.showCount !== false && (
                                        <span className="value-count">({value.count})</span>
                                    )}
                                </label>
                            ))
                        ) : (
                            <div className="no-values">
                                {searchTerm ? 'No matching filters found' : 'No filters available'}
                            </div>
                        )}
                    </div>

                </div>
            )}
        </div>
    );
};

/**
 * Main Custom Filter Component
 */
const CustomFilterComponent: React.FC<ICustomFilterComponentProps> = ({
    filters,
    onFilterUpdate,
    onOperatorUpdate,
    onApplyAll,
    onClearAll,
    theme = 'light',
    compact = false,
    showOperatorToggle = true
}) => {

    // Check if there are any selected filters
    const hasSelections = filters.some(filter => 
        filter.values.some(value => value.selected)
    );

    const handleFilterUpdate = useCallback((filterName: string, selectedValues: string[]) => {
        onFilterUpdate(filterName, selectedValues);
        
        // Dispatch custom event for PnP Modern Search integration
        const event = new CustomEvent(ExtensibilityConstants.EVENT_FILTER_UPDATED, {
            detail: {
                filterName,
                selectedValues
            }
        });
        document.dispatchEvent(event);
    }, [onFilterUpdate]);

    const handleOperatorUpdate = useCallback((filterName: string, operator: 'OR' | 'AND') => {
        onOperatorUpdate(filterName, operator);
        
        // Dispatch custom event for PnP Modern Search integration
        const event = new CustomEvent(ExtensibilityConstants.EVENT_FILTER_VALUE_OPERATOR_UPDATED, {
            detail: {
                filterName,
                operator
            }
        });
        document.dispatchEvent(event);
    }, [onOperatorUpdate]);

    const handleApplyAll = useCallback(() => {
        onApplyAll();
        
        // Dispatch custom event for PnP Modern Search integration
        const event = new CustomEvent(ExtensibilityConstants.EVENT_FILTER_APPLY_ALL, {
            detail: {}
        });
        document.dispatchEvent(event);
    }, [onApplyAll]);

    const handleClearAll = useCallback(() => {
        onClearAll();
        
        // Dispatch custom event for PnP Modern Search integration
        const event = new CustomEvent(ExtensibilityConstants.EVENT_FILTER_CLEAR_ALL, {
            detail: {}
        });
        document.dispatchEvent(event);
    }, [onClearAll]);

    if (!filters || filters.length === 0) {
        return (
            <div className={`custom-filter-component ${theme} empty`}>
                <div className="empty-state">
                    <i className="ms-Icon ms-Icon--Filter" aria-hidden="true" />
                    <p>No filters available</p>
                </div>
            </div>
        );
    }

    return (
        <div className={`custom-filter-component ${theme} ${compact ? 'compact' : ''}`}>
            
            {/* Filter Header */}
            <div className="filter-component-header">
                <h3 className="component-title">
                    <i className="ms-Icon ms-Icon--Filter" aria-hidden="true" />
                    Refine Results
                </h3>
                
                {hasSelections && (
                    <div className="global-actions">
                        <button
                            type="button"
                            className="global-action apply-all"
                            onClick={handleApplyAll}
                            title="Apply all selected filters"
                        >
                            <i className="ms-Icon ms-Icon--CheckMark" aria-hidden="true" />
                            Apply
                        </button>
                        <button
                            type="button"
                            className="global-action clear-all"
                            onClick={handleClearAll}
                            title="Clear all selected filters"
                        >
                            <i className="ms-Icon ms-Icon--Clear" aria-hidden="true" />
                            Clear All
                        </button>
                    </div>
                )}
            </div>

            {/* Filter Panels */}
            <div className="filter-panels">
                {filters.map((filter) => (
                    <FilterPanel
                        key={filter.filterName}
                        filter={filter}
                        onFilterUpdate={(selectedValues) => handleFilterUpdate(filter.filterName, selectedValues)}
                        onOperatorUpdate={(operator) => handleOperatorUpdate(filter.filterName, operator)}
                        theme={theme}
                        compact={compact}
                        showOperatorToggle={showOperatorToggle}
                    />
                ))}
            </div>

            {/* Status Bar */}
            {hasSelections && (
                <div className="filter-status">
                    <span className="status-text">
                        {filters.reduce((total, filter) => 
                            total + filter.values.filter(v => v.selected).length, 0
                        )} filters applied
                    </span>
                </div>
            )}

        </div>
    );
};

export default CustomFilterComponent;