# -*- coding: utf-8 -*-
# IronPython 2.7 Compatible
"""
HTML Report Generator Module

This module handles the generation of interactive HTML reports for view template comparisons,
including styling, JavaScript functionality, and organized data presentation.
"""

import os

from datetime import datetime

from EnneadTab import ERROR_HANDLE, FOLDER


class HTMLReportGenerator:
    """
    Generates interactive HTML reports for view template comparisons.
    
    This class is responsible for creating comprehensive, interactive HTML reports
    that display template differences in an organized and visually appealing format.
    """
    
    def __init__(self, template_names):
        """
        Initialize the HTML report generator.
        
        Args:
            template_names: List of template names for the report
        """
        self.template_names = template_names
    
    def generate_comparison_report(self, differences, summary_stats, comparison_data=None, json_file_path=None):
        """
        Generate the complete HTML comparison report.
        
        Args:
            differences: Dictionary containing all differences found
            summary_stats: Dictionary containing summary statistics
            comparison_data: Dictionary containing all template data for comprehensive comparison
            json_file_path: Path to the saved JSON file for clickable link
            
        Returns:
            str: Complete HTML report as string
        """
        # Ensure differences and summary_stats are valid dictionaries
        if not isinstance(differences, dict):
            differences = {}
        if not isinstance(summary_stats, dict):
            summary_stats = {}
        if comparison_data is None:
            comparison_data = {}
        
        # Light Unicode cleaning - only clean problematic surrogates, keep normal Unicode
        differences = self._clean_problematic_unicode(differences)
        summary_stats = self._clean_problematic_unicode(summary_stats) 
        comparison_data = self._clean_problematic_unicode(comparison_data)
        
        # Ensure template_names is properly initialized
        if not hasattr(self, 'template_names') or not isinstance(self.template_names, list):
            self.template_names = []
        
        # Validate template_names
        valid_template_names = []
        for name in self.template_names:
            try:
                if name is not None and str(name).strip():
                    valid_template_names.append(str(name).strip())
                else:
                    valid_template_names.append("Unknown Template")
            except Exception as e:
                ERROR_HANDLE.print_note("Error validating template name: {}".format(str(e)))
                valid_template_names.append("Error: Invalid Name")
        
        self.template_names = valid_template_names
        
        html = self._generate_html_header()
        html += self._generate_summary_section(summary_stats, json_file_path)
        
        # Only show detailed sections (differences) if we have differences
        if differences and isinstance(differences, dict) and any(len(section) > 0 for section in differences.values()):
            html += self._generate_detailed_sections(differences, comparison_data)
        
        html += self._generate_html_footer()
        
        return html
    
    def _generate_html_header(self):
        """
        Generate the HTML header with CSS and JavaScript.
        
        Returns:
            str: HTML header section
        """
        # Debug: Print template_names to see what we're working with
        ERROR_HANDLE.print_note("Template names: {}".format(self.template_names))
        
        # Ensure template_names is a valid list
        if not isinstance(self.template_names, list):
            self.template_names = []
        
        # Filter out any None or invalid template names
        valid_template_names = []
        for name in self.template_names:
            try:
                if name is not None and str(name).strip():
                    valid_template_names.append(str(name).strip())
                else:
                    valid_template_names.append("Unknown Template")
            except Exception as e:
                ERROR_HANDLE.print_note("Error processing template name: {}".format(str(e)))
                valid_template_names.append("Error: Invalid Name")
        
        # Use the validated template names
        self.template_names = valid_template_names
        
        # Debug: Print validated template names
        ERROR_HANDLE.print_note("Validated template names: {}".format(self.template_names))
        
        # Safely format the timestamp
        try:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        except Exception as e:
            timestamp = "Unknown time"
            ERROR_HANDLE.print_note("Error formatting timestamp: {}".format(str(e)))
        
        # Debug: Print timestamp
        ERROR_HANDLE.print_note("Timestamp: {}".format(timestamp))
        
        # Build HTML block by block
        html_parts = []
        
        # DOCTYPE and HTML start
        html_parts.append("<!DOCTYPE html>")
        html_parts.append("<html>")
        
        # Head section
        html_parts.append("<head>")
        html_parts.append("    <meta charset=\"UTF-8\">")
        html_parts.append("    <title>EnneadTab - View Template Comparison Report</title>")
        
        # CSS styles
        html_parts.append("    <style>")
        html_parts.append("        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');")
        html_parts.append("        body { font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; margin: 0; background: #1a1a1a; color: #ffffff; min-height: 100vh; line-height: 1.6; }")
        html_parts.append("        .container { max-width: 1400px; margin: 0 auto; background: #2a2a2a; padding: 30px; border-radius: 12px; box-shadow: 0 8px 32px rgba(0, 0, 0, 0.4); border: 1px solid #404040; }")
        html_parts.append("        h1 { color: #ffffff; text-align: center; border-bottom: 2px solid #666666; padding-bottom: 15px; font-size: 2.5em; font-weight: 600; margin-bottom: 30px; letter-spacing: -0.02em; }")
        html_parts.append("        h2 { color: #ffffff; cursor: pointer; padding: 15px; background: #333333; border-radius: 8px; margin: 15px 0; border: 1px solid #404040; transition: all 0.3s ease; font-weight: 500; }")
        html_parts.append("        h2:hover { background: #404040; transform: translateY(-1px); box-shadow: 0 4px 12px rgba(0, 0, 0, 0.3); }")
        html_parts.append("        .collapsible { display: none; padding: 20px; border: 1px solid #404040; border-radius: 8px; margin: 10px 0; background: #2a2a2a; }")
        html_parts.append("        table { width: 100%; border-collapse: collapse; margin: 15px 0; background: #2a2a2a; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 15px rgba(0, 0, 0, 0.3); }")
        html_parts.append("        th, td { border: 1px solid #404040; padding: 12px; text-align: left; }")
        html_parts.append("        th { background: #404040; color: #ffffff; font-weight: 600; font-size: 0.95em; }")
        html_parts.append("        .template-col { background: #333333; font-weight: 600; color: #cccccc; border-left: 3px solid #666666; }")
        html_parts.append("        .same { background: #2d4a2d; color: #a8d5a8; border: 1px solid #4a6b4a; }")
        html_parts.append("        .different { background: #4a3d2d; color: #d5c4a8; border: 1px solid #6b5a4a; }")
        html_parts.append("        .summary { background: #333333; padding: 25px; border-radius: 12px; margin: 20px 0; border: 1px solid #404040; }")
        html_parts.append("        .warning { background: #4a2d2d; border: 2px solid #6b4a4a; padding: 15px; margin-bottom: 20px; border-radius: 10px; color: #d5a8a8; }")
        html_parts.append("        .timestamp { color: #999999; font-size: 0.9em; text-align: center; margin: 15px 0; font-style: italic; }")
        html_parts.append("        .search-container { position: fixed; top: 0; left: 0; right: 0; background: #1a1a1a; padding: 20px; z-index: 1000; box-shadow: 0 4px 20px rgba(0, 0, 0, 0.5); border-bottom: 1px solid #404040; display: flex; justify-content: center; align-items: center; }")
        html_parts.append("        .search-container .search-content { display: flex; align-items: center; justify-content: center; max-width: 800px; width: 100%; flex-wrap: wrap; gap: 10px; }")
        html_parts.append("        .search-container input { width: 400px; min-width: 300px; padding: 12px 16px; border: 1px solid #404040; border-radius: 8px; font-size: 14px; margin-right: 12px; background: #2a2a2a; color: #ffffff; transition: all 0.3s ease; font-family: 'Inter', sans-serif; }")
        html_parts.append("        @media (max-width: 768px) {")
        html_parts.append("            .search-container .search-content { flex-direction: column; align-items: center; }")
        html_parts.append("            .search-container input { width: 90%; max-width: 400px; margin-right: 0; margin-bottom: 10px; }")
        html_parts.append("            .search-container .search-info { margin-left: 0; margin-top: 10px; text-align: center; }")
        html_parts.append("        }")
        html_parts.append("        .search-container input:focus { outline: none; border-color: #666666; box-shadow: 0 0 10px rgba(102, 102, 102, 0.3); background: #333333; }")
        html_parts.append("        .search-container input::placeholder { color: #999999; }")
        html_parts.append("        .search-container button { padding: 12px 20px; background: #404040; color: #ffffff; border: 1px solid #666666; border-radius: 8px; cursor: pointer; font-size: 14px; font-weight: 500; transition: all 0.3s ease; margin-right: 8px; font-family: 'Inter', sans-serif; }")
        html_parts.append("        .search-container button:hover { background: #666666; transform: translateY(-1px); box-shadow: 0 4px 12px rgba(0, 0, 0, 0.3); }")
        html_parts.append("        .search-container .search-info { color: #cccccc; font-size: 13px; margin-left: 15px; font-weight: 400; }")
        html_parts.append("        body { padding-top: 80px; }")
        html_parts.append("        .search-highlight { background: #4a4a2d; color: #d5d5a8; font-weight: 600; padding: 2px 4px; border-radius: 4px; }")
        html_parts.append("        .visible-cell { background-color: #2d4a2d !important; color: #a8d5a8 !important; font-weight: 600; text-align: center; border: 2px solid #4a6b4a; position: relative; }")
        html_parts.append("        .hidden-cell { background-color: #4a2d2d !important; color: #d5a8a8 !important; font-weight: 600; text-align: center; border: 2px solid #6b4a4a; position: relative; }")
        # Use Unicode escape sequences for IronPython 2.7 compatibility
        html_parts.append("        .visible-cell::before { content: '\\1F441 '; font-size: 14px; }")  # Eye emoji
        html_parts.append("        .hidden-cell::before { content: '\\1F6AB '; font-size: 14px; }")  # Prohibited emoji
        html_parts.append("        .search-hidden { display: none; }")
        html_parts.append("        ul { list-style: none; padding: 0; }")
        html_parts.append("        ul li { padding: 8px 0; border-bottom: 1px solid #404040; color: #ffffff; }")
        html_parts.append("        ul li:last-child { border-bottom: none; }")
        html_parts.append("        .error { background: #4a2d2d; border: 1px solid #6b4a4a; padding: 20px; border-radius: 10px; margin: 15px 0; color: #d5a8a8; }")
        html_parts.append("        .section { margin: 20px 0; }")
        html_parts.append("        .bold-text { font-weight: 700; color: #ffffff; }")
        html_parts.append("        /* Custom scrollbar */")
        html_parts.append("        ::-webkit-scrollbar { width: 12px; }")
        html_parts.append("        ::-webkit-scrollbar-track { background: #2a2a2a; border-radius: 6px; }")
        html_parts.append("        ::-webkit-scrollbar-thumb { background: #404040; border-radius: 6px; }")
        html_parts.append("        ::-webkit-scrollbar-thumb:hover { background: #666666; }")
        html_parts.append("    </style>")
        
        # JavaScript
        html_parts.append("    <script>")
        html_parts.append("        function toggleSection(sectionId) {")
        html_parts.append("            var section = document.getElementById(sectionId);")
        html_parts.append("            if (section.style.display === \"none\" || section.style.display === \"\") {")
        html_parts.append("                section.style.display = \"block\";")
        html_parts.append("            } else {")
        html_parts.append("                section.style.display = \"none\";")
        html_parts.append("            }")
        html_parts.append("        }")
        html_parts.append("        var searchTimeout;")
        html_parts.append("        function performSearch() {")
        html_parts.append("            var searchTerm = document.getElementById('searchInput').value.toLowerCase();")
        html_parts.append("            var searchInfo = document.getElementById('searchInfo');")
        html_parts.append("            var tables = document.querySelectorAll('table');")
        html_parts.append("            var totalRows = 0;")
        html_parts.append("            var matchingRows = 0;")
        html_parts.append("            document.querySelectorAll('.search-highlight').forEach(function(el) {")
        html_parts.append("                el.classList.remove('search-highlight');")
        html_parts.append("            });")
        html_parts.append("            document.querySelectorAll('.search-hidden').forEach(function(el) {")
        html_parts.append("                el.classList.remove('search-hidden');")
        html_parts.append("            });")
        html_parts.append("            if (searchTerm === '') {")
        html_parts.append("                searchInfo.textContent = 'Enter search term to filter results';")
        html_parts.append("                return;")
        html_parts.append("            }")
        html_parts.append("            tables.forEach(function(table) {")
        html_parts.append("                var rows = table.querySelectorAll('tr');")
        html_parts.append("                rows.forEach(function(row) {")
        html_parts.append("                    totalRows++;")
        html_parts.append("                    var cells = row.querySelectorAll('td, th');")
        html_parts.append("                    var rowText = '';")
        html_parts.append("                    var hasMatch = false;")
        html_parts.append("                    cells.forEach(function(cell) {")
        html_parts.append("                        rowText += cell.textContent + ' ';")
        html_parts.append("                    });")
        html_parts.append("                    if (rowText.toLowerCase().includes(searchTerm)) {")
        html_parts.append("                        hasMatch = true;")
        html_parts.append("                        matchingRows++;")
        html_parts.append("                        cells.forEach(function(cell) {")
        html_parts.append("                            var originalText = cell.innerHTML;")
        html_parts.append("                            var regex = new RegExp('(' + searchTerm + ')', 'gi');")
        html_parts.append("                            cell.innerHTML = originalText.replace(regex, '<span class=\"search-highlight\">$1</span>');")
        html_parts.append("                        });")
        html_parts.append("                    } else {")
        html_parts.append("                        if (row.querySelector('th')) {")
        html_parts.append("                        } else {")
        html_parts.append("                            row.classList.add('search-hidden');")
        html_parts.append("                        }")
        html_parts.append("                    }")
        html_parts.append("                });")
        html_parts.append("            });")
        html_parts.append("            searchInfo.textContent = 'Found ' + matchingRows + ' matching rows';")
        html_parts.append("        }")
        html_parts.append("        function debouncedSearch() {")
        html_parts.append("            clearTimeout(searchTimeout);")
        html_parts.append("            searchTimeout = setTimeout(performSearch, 300);")
        html_parts.append("        }")
        html_parts.append("        function handleKeyPress(event) {")
        html_parts.append("            if (event.key === 'Enter') {")
        html_parts.append("                clearTimeout(searchTimeout);")
        html_parts.append("                performSearch();")
        html_parts.append("            }")
        html_parts.append("        }")
        html_parts.append("        function clearSearch() {")
        html_parts.append("            document.getElementById('searchInput').value = '';")
        html_parts.append("            performSearch();")
        html_parts.append("        }")
        html_parts.append("    </script>")
        html_parts.append("</head>")
        
        # Body start
        html_parts.append("<body>")
        
        # Search bar
        html_parts.append("    <div class=\"search-container\">")
        html_parts.append("        <div class=\"search-content\">")
        html_parts.append("            <input type=\"text\" id=\"searchInput\" placeholder=\"Search for categories, parameters, filters...\" oninput=\"debouncedSearch()\" onkeypress=\"handleKeyPress(event)\">")
        html_parts.append("            <button onclick=\"clearSearch()\">Clear</button>")
        html_parts.append("            <span class=\"search-info\" id=\"searchInfo\">Enter search term to filter results</span>")
        html_parts.append("        </div>")
        html_parts.append("    </div>")
        
        # Container start
        html_parts.append("    <div class=\"container\">")
        html_parts.append("        <h1>EnneadTab - View Template Comparison Report</h1>")
        html_parts.append("        <div style=\"text-align: center; color: #cccccc; font-size: 1.1em; margin-bottom: 20px; font-weight: 400;\">Powered by EnneadTab Ecosystem</div>")
        html_parts.append("        <div class=\"timestamp\">Generated on: " + timestamp + "</div>")
        html_parts.append("        <div class=\"summary\">")
        html_parts.append("            <h3>Templates Compared:</h3>")
        html_parts.append("            <ul>")
        
        # Template names list
        for template_name in self.template_names:
            try:
                if template_name and str(template_name).strip():
                    html_parts.append("                <li>" + str(template_name).strip() + "</li>")
                else:
                    html_parts.append("                <li>Unknown Template</li>")
            except Exception as e:
                ERROR_HANDLE.print_note("Error processing template name in header: {}".format(str(e)))
                html_parts.append("                <li>Error: Invalid Template Name</li>")
        
        # Close template list and summary
        html_parts.append("            </ul>")
        html_parts.append("        </div>")
        
        return "\n".join(html_parts)
    
    def _generate_comprehensive_comparison_section(self, comparison_data):
        """
        Generate a comprehensive side-by-side comparison of ALL settings.
        
        Args:
            comparison_data: Dictionary containing all template data
            
        Returns:
            str: HTML for comprehensive comparison section
        """
        html_parts = []
        
        html_parts.append("        <div class=\"summary\">")
        html_parts.append("            <h3>Comprehensive Comparison - All Settings</h3>")
        html_parts.append("            <p>Below is a complete side-by-side comparison of all settings across the selected templates:</p>")
        html_parts.append("        </div>")
        
        # Category Overrides Comprehensive Table
        html_parts.append("        <div class=\"section\">")
        html_parts.append("            <h2 class=\"toggle\" onclick=\"toggleSection('comprehensive_categories')\">All Category Overrides (Click to expand)</h2>")
        html_parts.append("            <div id=\"comprehensive_categories\" class=\"collapsible\">")
        html_parts.append("                <table>")
        html_parts.append("                    <tr>")
        html_parts.append("                        <th>Category</th>")
        
        for template_name in self.template_names:
            html_parts.append("                        <th>" + template_name + "</th>")
        
        html_parts.append("                    </tr>")
        
        # Get all categories from all templates
        all_categories = set()
        for template_name, data in comparison_data.items():
            if 'category_overrides' in data:
                all_categories.update(data['category_overrides'].keys())
        
        for category in sorted(all_categories):
            html_parts.append("                    <tr>")
            html_parts.append("                        <td class=\"template-col\">" + category + "</td>")
            
            for template_name in self.template_names:
                try:
                    template_data = comparison_data.get(template_name, {})
                    category_overrides = template_data.get('category_overrides', {})
                    override_data = category_overrides.get(category, None)
                    
                    if override_data is None:
                        html_parts.append("                        <td class=\"same\">No Override</td>")
                    elif override_data == "UNCONTROLLED":
                        html_parts.append("                        <td style=\"background: linear-gradient(135deg, rgba(255, 165, 2, 0.3) 0%, rgba(255, 165, 2, 0.2) 100%); color: #ffa502; font-weight: bold;\">UNCONTROLLED</td>")
                    else:
                        summary = self._create_override_summary(override_data)
                        html_parts.append("                        <td class=\"different\" style=\"white-space: pre-line;\">" + summary + "</td>")
                except Exception as e:
                    ERROR_HANDLE.print_note("Error processing category override for {} in {}: {}".format(category, template_name, str(e)))
                    html_parts.append("                        <td class=\"error\">Error</td>")
            
            html_parts.append("                    </tr>")
        
        html_parts.append("                </table>")
        html_parts.append("            </div>")
        html_parts.append("        </div>")
        
        # Category Visibility Comprehensive Table
        html_parts.append("        <div class=\"section\">")
        html_parts.append("            <h2 class=\"toggle\" onclick=\"toggleSection('comprehensive_visibility')\">All Category Visibility (Click to expand)</h2>")
        html_parts.append("            <div id=\"comprehensive_visibility\" class=\"collapsible\">")
        html_parts.append("                <table>")
        html_parts.append("                    <tr>")
        html_parts.append("                        <th>Category</th>")
        
        for template_name in self.template_names:
            html_parts.append("                        <th>" + template_name + "</th>")
        
        html_parts.append("                    </tr>")
        
        # Get all categories from all templates
        all_categories = set()
        for template_name, data in comparison_data.items():
            if 'category_visibility' in data:
                all_categories.update(data['category_visibility'].keys())
        
        for category in sorted(all_categories):
            html_parts.append("                    <tr>")
            html_parts.append("                        <td class=\"template-col\">" + category + "</td>")
            
            for template_name in self.template_names:
                try:
                    template_data = comparison_data.get(template_name, {})
                    category_visibility = template_data.get('category_visibility', {})
                    visibility = category_visibility.get(category, 'N/A')
                    
                    if visibility == "UNCONTROLLED":
                        html_parts.append("                        <td style=\"background: linear-gradient(135deg, rgba(255, 165, 2, 0.3) 0%, rgba(255, 165, 2, 0.2) 100%); color: #ffa502; font-weight: bold;\">UNCONTROLLED</td>")
                    elif visibility in ['On', 'Visible']:
                        html_parts.append("                        <td class=\"visible-cell\">" + visibility + "</td>")
                    elif visibility == 'Hidden':
                        html_parts.append("                        <td class=\"hidden-cell\">" + visibility + "</td>")
                    else:
                        html_parts.append("                        <td class=\"different\">" + str(visibility) + "</td>")
                except Exception as e:
                    ERROR_HANDLE.print_note("Error processing category visibility for {} in {}: {}".format(category, template_name, str(e)))
                    html_parts.append("                        <td class=\"error\">Error</td>")
            
            html_parts.append("                    </tr>")
        
        html_parts.append("                </table>")
        html_parts.append("            </div>")
        html_parts.append("        </div>")
        
        # Workset Visibility Comprehensive Table
        html_parts.append("        <div class=\"section\">")
        html_parts.append("            <h2 class=\"toggle\" onclick=\"toggleSection('comprehensive_worksets')\">All Workset Visibility (Click to expand)</h2>")
        html_parts.append("            <div id=\"comprehensive_worksets\" class=\"collapsible\">")
        html_parts.append("                <table>")
        html_parts.append("                    <tr>")
        html_parts.append("                        <th>Workset</th>")
        
        for template_name in self.template_names:
            html_parts.append("                        <th>" + template_name + "</th>")
        
        html_parts.append("                    </tr>")
        
        # Get all worksets from all templates
        all_worksets = set()
        for template_name, data in comparison_data.items():
            if 'workset_visibility' in data:
                all_worksets.update(data['workset_visibility'].keys())
        
        for workset in sorted(all_worksets):
            html_parts.append("                    <tr>")
            html_parts.append("                        <td class=\"template-col\">" + workset + "</td>")
            
            for template_name in self.template_names:
                try:
                    template_data = comparison_data.get(template_name, {})
                    workset_visibility = template_data.get('workset_visibility', {})
                    visibility = workset_visibility.get(workset, 'N/A')
                    
                    if visibility == "UNCONTROLLED":
                        html_parts.append("                        <td style=\"background: linear-gradient(135deg, rgba(255, 165, 2, 0.3) 0%, rgba(255, 165, 2, 0.2) 100%); color: #ffa502; font-weight: bold;\">UNCONTROLLED</td>")
                    elif visibility in ['On', 'Visible']:
                        html_parts.append("                        <td class=\"same\">" + visibility + "</td>")
                    elif visibility == 'Hidden':
                        html_parts.append("                        <td class=\"different\">" + visibility + "</td>")
                    else:
                        html_parts.append("                        <td class=\"different\">" + str(visibility) + "</td>")
                except Exception as e:
                    ERROR_HANDLE.print_note("Error processing workset visibility for {} in {}: {}".format(workset, template_name, str(e)))
                    html_parts.append("                        <td class=\"error\">Error</td>")
            
            html_parts.append("                    </tr>")
        
        html_parts.append("                </table>")
        html_parts.append("            </div>")
        html_parts.append("        </div>")
        
        # View Parameters Comprehensive Table
        html_parts.append("        <div class=\"section\">")
        html_parts.append("            <h2 class=\"toggle\" onclick=\"toggleSection('comprehensive_parameters')\">All View Parameters (Click to expand)</h2>")
        html_parts.append("            <div id=\"comprehensive_parameters\" class=\"collapsible\">")
        html_parts.append("                <table>")
        html_parts.append("                    <tr>")
        html_parts.append("                        <th>Parameter</th>")
        
        for template_name in self.template_names:
            html_parts.append("                        <th>" + template_name + "</th>")
        
        html_parts.append("                    </tr>")
        
        # Get all parameters from all templates
        all_parameters = set()
        for template_name, data in comparison_data.items():
            if 'view_parameters' in data:
                all_parameters.update(data['view_parameters'])
        
        for parameter in sorted(all_parameters):
            html_parts.append("                    <tr>")
            html_parts.append("                        <td class=\"template-col\">" + parameter + "</td>")
            
            for template_name in self.template_names:
                try:
                    template_data = comparison_data.get(template_name, {})
                    view_parameters = template_data.get('view_parameters', {})  # Now a dict
                    uncontrolled_parameters = template_data.get('uncontrolled_parameters', [])
                    
                    if parameter in view_parameters:
                        # Parameter is controlled, show its value
                        param_value = view_parameters[parameter]
                        html_parts.append("                        <td style=\"background-color: #d1ecf1; color: #0c5460; font-weight: bold; text-align: center;\">{}</td>".format(str(param_value)))
                    elif parameter in uncontrolled_parameters:
                        # Parameter is not controlled
                        html_parts.append("                        <td style=\"background-color: #fff3cd; color: #856404; font-weight: bold; text-align: center;\">Not Controlled</td>")
                    else:
                        # Parameter not present
                        html_parts.append("                        <td style=\"background-color: #f8f9fa; color: #6c757d; font-style: italic; text-align: center;\">Not Present</td>")
                except Exception as e:
                    ERROR_HANDLE.print_note("Error processing view parameter {} in {}: {}".format(parameter, template_name, str(e)))
                    html_parts.append("                        <td class=\"error\">Error</td>")
            
            html_parts.append("                    </tr>")
        
        html_parts.append("                </table>")
        html_parts.append("            </div>")
        html_parts.append("        </div>")
        
        # Uncontrolled Parameters Comprehensive Table
        html_parts.append("        <div class=\"section\">")
        html_parts.append("            <h2 class=\"toggle\" onclick=\"toggleSection('comprehensive_uncontrolled')\" style=\"color: #ff6b6b;\">All Uncontrolled Parameters - DANGEROUS (Click to expand)</h2>")
        html_parts.append("            <div id=\"comprehensive_uncontrolled\" class=\"collapsible\">")
        html_parts.append("                <div class=\"warning\">")
        html_parts.append("                    <strong>WARNING:</strong> Uncontrolled parameters are NOT marked as included when the view is used as a template.")
        html_parts.append("                    This can cause inconsistent behavior across views using the same template.")
        html_parts.append("                </div>")
        html_parts.append("                <table>")
        html_parts.append("                    <tr>")
        html_parts.append("                        <th>Parameter</th>")
        
        for template_name in self.template_names:
            html_parts.append("                        <th>" + template_name + "</th>")
        
        html_parts.append("                    </tr>")
        
        # Get all uncontrolled parameters from all templates
        all_uncontrolled = set()
        for template_name, data in comparison_data.items():
            if 'uncontrolled_parameters' in data:
                all_uncontrolled.update(data['uncontrolled_parameters'])
        
        for parameter in sorted(all_uncontrolled):
            html_parts.append("                    <tr>")
            html_parts.append("                        <td class=\"template-col\">" + parameter + "</td>")
            
            for template_name in self.template_names:
                try:
                    template_data = comparison_data.get(template_name, {})
                    uncontrolled_parameters = template_data.get('uncontrolled_parameters', [])
                    is_uncontrolled = parameter in uncontrolled_parameters
                    
                    if is_uncontrolled:
                        html_parts.append("                        <td style=\"background: linear-gradient(135deg, rgba(255, 107, 107, 0.3) 0%, rgba(255, 107, 107, 0.2) 100%); color: #ff6b6b; font-weight: bold;\">Uncontrolled</td>")
                    else:
                        html_parts.append("                        <td class=\"same\">Controlled</td>")
                except Exception as e:
                    ERROR_HANDLE.print_note("Error processing uncontrolled parameter {} in {}: {}".format(parameter, template_name, str(e)))
                    html_parts.append("                        <td class=\"error\">Error</td>")
            
            html_parts.append("                    </tr>")
        
        html_parts.append("                </table>")
        html_parts.append("            </div>")
        html_parts.append("        </div>")
        
        # Filters Comprehensive Table
        html_parts.append("        <div class=\"section\">")
        html_parts.append("            <h2 class=\"toggle\" onclick=\"toggleSection('comprehensive_filters')\">All Filters (Click to expand)</h2>")
        html_parts.append("            <div id=\"comprehensive_filters\" class=\"collapsible\">")
        html_parts.append("                <table>")
        html_parts.append("                    <tr>")
        html_parts.append("                        <th>Filter</th>")
        
        for template_name in self.template_names:
            html_parts.append("                        <th>" + template_name + "</th>")
        
        html_parts.append("                    </tr>")
        
        # Get all filters from all templates
        all_filters = set()
        for template_name, data in comparison_data.items():
            if 'filters' in data:
                all_filters.update(data['filters'].keys())
        
        for filter_name in sorted(all_filters):
            html_parts.append("                    <tr>")
            html_parts.append("                        <td class=\"template-col\">" + filter_name + "</td>")
            
            for template_name in self.template_names:
                try:
                    template_data = comparison_data.get(template_name, {})
                    filters = template_data.get('filters', {})
                    filter_data = filters.get(filter_name, None)
                    
                    if filter_data is None:
                        html_parts.append("                        <td class=\"same\">Not Applied</td>")
                    elif filter_data == "UNCONTROLLED":
                        html_parts.append("                        <td style=\"background: linear-gradient(135deg, rgba(255, 165, 2, 0.3) 0%, rgba(255, 165, 2, 0.2) 100%); color: #ffa502; font-weight: bold;\">UNCONTROLLED</td>")
                    else:
                        summary = self._create_override_summary(filter_data)
                        html_parts.append("                        <td class=\"different\" style=\"white-space: pre-line;\">" + summary + "</td>")
                except Exception as e:
                    ERROR_HANDLE.print_note("Error processing filter {} in {}: {}".format(filter_name, template_name, str(e)))
                    html_parts.append("                        <td class=\"error\">Error</td>")
            
            html_parts.append("                    </tr>")
        
        html_parts.append("                </table>")
        html_parts.append("            </div>")
        html_parts.append("        </div>")
        
        # Import Categories Comprehensive Table
        html_parts.append("        <div class=\"section\">")
        html_parts.append("            <h2 class=\"toggle\" onclick=\"toggleSection('comprehensive_imports')\">All Import Categories (Click to expand)</h2>")
        html_parts.append("            <div id=\"comprehensive_imports\" class=\"collapsible\">")
        html_parts.append("                <table>")
        html_parts.append("                    <tr>")
        html_parts.append("                        <th>Import Category</th>")
        
        for template_name in self.template_names:
            html_parts.append("                        <th>" + template_name + "</th>")
        
        html_parts.append("                    </tr>")
        
        # Get all import categories from all templates
        all_imports = set()
        for template_name, data in comparison_data.items():
            if 'import_categories' in data:
                all_imports.update(data['import_categories'].keys())
        
        for import_name in sorted(all_imports):
            html_parts.append("                    <tr>")
            html_parts.append("                        <td class=\"template-col\">" + import_name + "</td>")
            
            for template_name in self.template_names:
                try:
                    template_data = comparison_data.get(template_name, {})
                    import_categories = template_data.get('import_categories', {})
                    import_data = import_categories.get(import_name, None)
                    
                    if import_data is None:
                        html_parts.append("                        <td class=\"same\">No Override</td>")
                    elif import_data == "UNCONTROLLED":
                        html_parts.append("                        <td style=\"background: linear-gradient(135deg, rgba(255, 165, 2, 0.3) 0%, rgba(255, 165, 2, 0.2) 100%); color: #ffa502; font-weight: bold;\">UNCONTROLLED</td>")
                    else:
                        summary = self._create_override_summary(import_data)
                        html_parts.append("                        <td class=\"different\" style=\"white-space: pre-line;\">" + summary + "</td>")
                except Exception as e:
                    ERROR_HANDLE.print_note("Error processing import category {} in {}: {}".format(import_name, template_name, str(e)))
                    html_parts.append("                        <td class=\"error\">Error</td>")
            
            html_parts.append("                    </tr>")
        
        html_parts.append("                </table>")
        html_parts.append("            </div>")
        html_parts.append("        </div>")
        
        # Revit Links Comprehensive Table
        html_parts.append("        <div class=\"section\">")
        html_parts.append("            <h2 class=\"toggle\" onclick=\"toggleSection('comprehensive_links')\">All Revit Links (Click to expand)</h2>")
        html_parts.append("            <div id=\"comprehensive_links\" class=\"collapsible\">")
        html_parts.append("                <table>")
        html_parts.append("                    <tr>")
        html_parts.append("                        <th>Revit Link</th>")
        
        for template_name in self.template_names:
            html_parts.append("                        <th>" + template_name + "</th>")
        
        html_parts.append("                    </tr>")
        
        # Get all Revit links from all templates
        all_links = set()
        for template_name, data in comparison_data.items():
            if 'revit_links' in data:
                all_links.update(data['revit_links'].keys())
        
        for link_name in sorted(all_links):
            html_parts.append("                    <tr>")
            html_parts.append("                        <td class=\"template-col\">" + link_name + "</td>")
            
            for template_name in self.template_names:
                try:
                    template_data = comparison_data.get(template_name, {})
                    revit_links = template_data.get('revit_links', {})
                    link_data = revit_links.get(link_name, {})
                    
                    if not link_data:
                        html_parts.append("                        <td class=\"same\">No Override</td>")
                    else:
                        # Create summary of link settings
                        settings = []
                        if 'visibility' in link_data:
                            settings.append("Visibility: " + str(link_data['visibility']))
                        if 'halftone' in link_data and link_data['halftone']:
                            settings.append("Halftone")
                        if 'underlay' in link_data and link_data['underlay']:
                            settings.append("Underlay")
                        if 'display_settings' in link_data:
                            settings.append("Display: " + str(link_data['display_settings']))
                        
                        if settings:
                            html_parts.append("                        <td class=\"different\">" + "; ".join(settings) + "</td>")
                        else:
                            html_parts.append("                        <td class=\"same\">Default</td>")
                except Exception as e:
                    ERROR_HANDLE.print_note("Error processing Revit link {} in {}: {}".format(link_name, template_name, str(e)))
                    html_parts.append("                        <td class=\"error\">Error</td>")
            
            html_parts.append("                    </tr>")
        
        html_parts.append("                </table>")
        html_parts.append("            </div>")
        html_parts.append("        </div>")
        
        # Detail Levels Comprehensive Table
        html_parts.append("        <div class=\"section\">")
        html_parts.append("            <h2 class=\"toggle\" onclick=\"toggleSection('comprehensive_detail_levels')\">All Detail Levels (Click to expand)</h2>")
        html_parts.append("            <div id=\"comprehensive_detail_levels\" class=\"collapsible\">")
        html_parts.append("                <table>")
        html_parts.append("                    <tr>")
        html_parts.append("                        <th>Category</th>")
        
        for template_name in self.template_names:
            html_parts.append("                        <th>" + template_name + "</th>")
        
        html_parts.append("                    </tr>")
        
        # Get all detail levels from all templates
        all_detail_levels = set()
        for template_name, data in comparison_data.items():
            if 'detail_levels' in data:
                all_detail_levels.update(data['detail_levels'].keys())
        
        for category_name in sorted(all_detail_levels):
            html_parts.append("                    <tr>")
            html_parts.append("                        <td class=\"template-col\">" + category_name + "</td>")
            
            for template_name in self.template_names:
                try:
                    template_data = comparison_data.get(template_name, {})
                    detail_levels = template_data.get('detail_levels', {})
                    detail_level = detail_levels.get(category_name, "By View")
                    
                    if detail_level == "By View":
                        html_parts.append("                        <td class=\"same\">" + detail_level + "</td>")
                    else:
                        html_parts.append("                        <td class=\"different\">" + detail_level + "</td>")
                except Exception as e:
                    ERROR_HANDLE.print_note("Error processing detail level for {} in {}: {}".format(category_name, template_name, str(e)))
                    html_parts.append("                        <td class=\"error\">Error</td>")
            
            html_parts.append("                    </tr>")
        
        html_parts.append("                </table>")
        html_parts.append("            </div>")
        html_parts.append("        </div>")
        
        # Linetypes Comprehensive Table
        html_parts.append("        <div class=\"section\">")
        html_parts.append("            <h2 class=\"toggle\" onclick=\"toggleSection('comprehensive_linetypes')\">All Linetypes (Click to expand)</h2>")
        html_parts.append("            <div id=\"comprehensive_linetypes\" class=\"collapsible\">")
        html_parts.append("                <table>")
        html_parts.append("                    <tr>")
        html_parts.append("                        <th>Linetype</th>")
        
        for template_name in self.template_names:
            html_parts.append("                        <th>" + template_name + "</th>")
        
        html_parts.append("                    </tr>")
        
        # Get all linetypes from all templates
        all_linetypes = set()
        for template_name, data in comparison_data.items():
            if 'linetypes' in data:
                all_linetypes.update(data['linetypes'].keys())
        
        for linetype_name in sorted(all_linetypes):
            html_parts.append("                    <tr>")
            html_parts.append("                        <td class=\"template-col\">" + linetype_name + "</td>")
            
            for template_name in self.template_names:
                try:
                    template_data = comparison_data.get(template_name, {})
                    linetypes = template_data.get('linetypes', {})
                    linetype_data = linetypes.get(linetype_name, {})
                    
                    if not linetype_data:
                        html_parts.append("                        <td class=\"same\">Not Available</td>")
                    else:
                        # Show linetype information
                        info = []
                        if 'id' in linetype_data:
                            info.append("ID: " + str(linetype_data['id']))
                        if 'is_used' in linetype_data:
                            info.append("Used: " + str(linetype_data['is_used']))
                        
                        if info:
                            html_parts.append("                        <td class=\"different\">" + "; ".join(info) + "</td>")
                        else:
                            html_parts.append("                        <td class=\"same\">Available</td>")
                except Exception as e:
                    ERROR_HANDLE.print_note("Error processing linetype {} in {}: {}".format(linetype_name, template_name, str(e)))
                    html_parts.append("                        <td class=\"error\">Error</td>")
            
            html_parts.append("                    </tr>")
        
        html_parts.append("                </table>")
        html_parts.append("            </div>")
        html_parts.append("        </div>")
        
        return "\n".join(html_parts)
    
    def _generate_summary_section(self, summary_stats, json_file_path=None):
        """
        Generate the summary section with statistics.
        
        Args:
            summary_stats: Dictionary containing summary statistics
            json_file_path: Path to the saved JSON file for clickable link
            
        Returns:
            str: HTML summary section
        """
        # Safely get values with defaults to prevent KeyError
        total_differences = summary_stats.get('total_differences', 0)
        category_overrides = summary_stats.get('category_overrides', 0)
        category_visibility = summary_stats.get('category_visibility', 0)
        workset_visibility = summary_stats.get('workset_visibility', 0)
        view_parameters = summary_stats.get('view_parameters', 0)
        uncontrolled_parameters = summary_stats.get('uncontrolled_parameters', 0)
        filters = summary_stats.get('filters', 0)
        import_categories = summary_stats.get('import_categories', 0)
        revit_links = summary_stats.get('revit_links', 0)
        detail_levels = summary_stats.get('detail_levels', 0)
        view_properties = summary_stats.get('view_properties', 0)
        
        # Get detailed category visibility breakdown
        category_visibility_breakdown = summary_stats.get('category_visibility_breakdown', {})
        
        # Create file path display for JSON file if path is provided
        json_path_html = ""
        if json_file_path:
            try:
                import os
                absolute_path = os.path.abspath(json_file_path)
                json_path_html = '<strong>JSON File Path:</strong> <span style="font-family: \'JetBrains Mono\', monospace; color: #cccccc; background: #2a2a2a; padding: 4px 8px; border-radius: 4px; border: 1px solid #404040;">{}</span>'.format(absolute_path)
            except Exception as e:
                ERROR_HANDLE.print_note("Error creating JSON file path display: {}".format(str(e)))
                json_path_html = ""
        
        html = """
        <div class="summary">
            <h3>Summary</h3>
            <p><strong>Total differences found:</strong> {}</p>
            <ul>
                <li>Category Overrides: {}</li>
                <li>Category Visibility: {}</li>
                <li>Workset Visibility: {}</li>
                <li>View Parameters: {}</li>
                <li>Uncontrolled Parameters: {} <span style="color: red; font-weight: bold;">DANGEROUS</span></li>
                <li>Filters: {}</li>
                <li>Import Categories: {}</li>
                <li>Revit Links: {}</li>
                <li>Detail Levels: {}</li>
                <li>View Properties: {}</li>
            </ul>
            <div style="margin-top: 15px; padding: 10px; background-color: #333333; border-left: 4px solid #666666; border-radius: 3px;">
                <p style="margin: 0; font-size: 0.9em; color: #cccccc;">
                    <strong>Note:</strong> Complete template data has been saved as a JSON file in the EnneadTab dump folder.
                </p>
                <div style="margin-top: 10px;">
                    {}
                </div>
                <p style="margin: 10px 0 0 0; font-size: 0.9em; color: #999999;">
                    Check the JSON file for detailed settings of all categories, filters, and parameters.
                </p>
        </div>
        </div>
""".format(json_path_html, total_differences, category_overrides, category_visibility, 
           workset_visibility, view_parameters, uncontrolled_parameters, filters,
           import_categories, revit_links, detail_levels, view_properties)
        
        # Add detailed category visibility breakdown if available
        if category_visibility_breakdown and isinstance(category_visibility_breakdown, dict):
            html += """
        <div class="summary">
            <h3>Category Visibility Breakdown</h3>
            <table style="width: 100%; margin: 10px 0;">
                <tr>
                    <th>Template</th>
                    <th>On/Visible</th>
                    <th>Hidden</th>
                    <th>Uncontrolled</th>
                    <th>Total Categories</th>
                </tr>
"""
            
            for template_name in self.template_names:
                try:
                    template_data = category_visibility_breakdown.get(template_name, {})
                    on_visible = template_data.get('on_visible', 0)
                    hidden = template_data.get('hidden', 0)
                    uncontrolled = template_data.get('uncontrolled', 0)
                    total = on_visible + hidden + uncontrolled
                    
                    html += """
                <tr>
                    <td><strong>{}</strong></td>
                    <td style="background-color: #d4edda; color: #155724; text-align: center;">{}</td>
                    <td style="background-color: #f8d7da; color: #721c24; text-align: center;">{}</td>
                    <td style="background-color: #fff3cd; color: #856404; text-align: center;">{}</td>
                    <td style="background-color: #e7f3ff; text-align: center;"><strong>{}</strong></td>
                </tr>
""".format(template_name, on_visible, hidden, uncontrolled, total)
                except Exception as e:
                    ERROR_HANDLE.print_note("Error processing category visibility breakdown for template {}: {}".format(template_name, str(e)))
                    html += """
                <tr>
                    <td><strong>{}</strong></td>
                    <td colspan="4" style="color: red;">Error processing data</td>
                </tr>
""".format(template_name)
            
            html += """
            </table>
            <div style="margin-top: 10px; font-size: 0.9em; color: #666;">
                <strong>Legend:</strong> 
                <span style="background-color: #d4edda; color: #155724; padding: 2px 6px; border-radius: 3px;">On/Visible</span> - Categories that are visible in the view
                <span style="background-color: #f8d7da; color: #721c24; padding: 2px 6px; border-radius: 3px; margin-left: 10px;">Hidden</span> - Categories that are hidden in the view
                <span style="background-color: #fff3cd; color: #856404; padding: 2px 6px; border-radius: 3px; margin-left: 10px;">Uncontrolled</span> - Categories not controlled by the template
            </div>
        </div>
"""
        
        if total_differences == 0:
            html += """
        <div class="summary" style="background-color: #d4edda; border: 1px solid #c3e6cb;">
            <h3>No Differences Found</h3>
            <p>All selected view templates have identical settings.</p>
        </div>
"""
        
        return html
    
    def _generate_detailed_sections(self, differences, comparison_data=None):
        """
        Generate all detailed comparison sections.
        
        Args:
            differences: Dictionary containing all differences
            comparison_data: Dictionary containing template data with usage information
            
        Returns:
            str: HTML for all detailed sections
        """
        html = ""
        
        try:
            # Safely access differences dictionary with defaults
            category_overrides = differences.get('category_overrides', {})
            category_visibility = differences.get('category_visibility', {})
            workset_visibility = differences.get('workset_visibility', {})
            view_parameters = differences.get('view_parameters', {})
            uncontrolled_parameters = differences.get('uncontrolled_parameters', {})
            filters = differences.get('filters', {})
        
        # Category Overrides Section
            if category_overrides:
                html += self._generate_category_overrides_section(category_overrides)
        
        # Category Visibility Section
            if category_visibility:
                html += self._generate_category_visibility_section(category_visibility)
        
        # Workset Visibility Section
            if workset_visibility:
                html += self._generate_workset_visibility_section(workset_visibility)
        
        # View Parameters Section
            if view_parameters:
                html += self._generate_view_parameters_section(view_parameters)
        
        # Uncontrolled Parameters Section (DANGEROUS)
            if uncontrolled_parameters:
                html += self._generate_uncontrolled_parameters_section(uncontrolled_parameters)
        
        # Filters Section
            if filters:
                html += self._generate_filters_section(filters)
        
        # Template Usage Section
            html += self._generate_template_usage_section(comparison_data)
                
        except Exception as e:
            # Add error section if detailed sections generation fails
            html += """
        <div class="section">
            <h2 style="color: red;">Error Generating Detailed Sections</h2>
            <div class="error">
                <p>An error occurred while generating detailed comparison sections:</p>
                <p><strong>{}</strong></p>
            </div>
        </div>
""".format(str(e))
        
        return html
    
    def _generate_category_overrides_section(self, differences):
        """
        Generate the category overrides comparison section.
        
        Args:
            differences: Dictionary of category override differences
            
        Returns:
            str: HTML for category overrides section
        """
        try:
            html = """
        <div class="section">
            <h2 class="toggle" onclick="toggleSection('category_overrides')">Category Overrides (Click to expand)</h2>
            <div id="category_overrides" class="collapsible">
                <table>
                    <tr>
                        <th>Category</th>
"""
            
            for template_name in self.template_names:
                html += "                        <th>{}</th>\n".format(template_name)
            
            html += "                    </tr>\n"
            
            # Render rows directly from differences (engine already filtered to only-different)
            if differences:
                for category, values in differences.items():
                    html += "                    <tr><td class='template-col'>{}</td>".format(category)
                    # Build a lookup of all template values for diff summary
                    for template_name in self.template_names:
                        value = values.get(template_name, None)
                        if value == "UNCONTROLLED":
                            html += "<td style='background-color: #FF8C00; color: white; font-weight: bold;'>UNCONTROLLED</td>"
                        elif value:
                            # Show the actual override summary for this template
                            summary = self._create_override_summary(value)
                            if summary:
                                # Process the summary to make section headers bold
                                processed_summary = self._process_summary_for_bold_headers(summary)
                                html += "<td class='different' style='white-space: pre-line;'>{}</td>".format(processed_summary)
                            else:
                                html += "<td class='same'>No Override</td>"
                        else:
                            # No local override
                            html += "<td class='same'></td>"
                    
                    html += "</tr>\n"
            else:
                # No categories with differences were found by the engine
                html += """
                    <tr>
                        <td colspan=\"{}\" style=\"text-align: center; color: #999999; font-style: italic; padding: 20px;\">
                            All templates have identical category override settings. No differences found.
                        </td>
                    </tr>
""".format(len(self.template_names) + 1)
            
            html += "                </table></div></div>\n"
            return html
            
        except Exception as e:
            # Return error section if generation fails
            return """
        <div class="section">
            <h2 class="toggle" onclick="toggleSection('category_overrides')">Category Overrides (Click to expand)</h2>
            <div id="category_overrides" class="collapsible">
                <div class="error">
                    <p>Error generating category overrides section: {}</p>
                </div>
            </div>
        </div>
""".format(str(e))
    
    def _create_override_summary(self, override_data):
        """
        Create a summary string for override data matching Revit UI structure.
        
        Args:
            override_data: Dictionary containing override details
            
        Returns:
            str: Summary string
        """
        if not override_data or override_data == "UNCONTROLLED":
            return str(override_data)
            
        summary_parts = []
        
        # Visibility
        if 'visibility' in override_data and override_data['visibility'] not in ['Visible', 'On']:
            summary_parts.append("Visibility: " + str(override_data['visibility']))
        
        # Projection/Surface - Lines
        line_parts = []
        if override_data.get('projection_line_weight'):
            weight_value = override_data['projection_line_weight']
            # Line weight -1 means using object style definition (no local override)
            if weight_value == -1:
                # line_parts.append("Line Weight: Default")
                pass
            else:
                line_parts.append("Line Weight: " + str(weight_value))
        if override_data.get('projection_line_color') and override_data['projection_line_color'] != 'Default':
            line_parts.append("Line Color: " + str(override_data['projection_line_color']))
        if override_data.get('projection_line_pattern') and override_data['projection_line_pattern'] != 'Default':
            line_parts.append("Line Pattern: " + str(override_data['projection_line_pattern']))
        if line_parts:
            summary_parts.append("Projection Lines: " + ", ".join(line_parts))
        
        # Projection/Surface - Patterns (Surface Foreground and Background)
        pattern_parts = []
        if override_data.get('surface_foreground_pattern') and override_data['surface_foreground_pattern'] != 'Default':
            pattern_parts.append("Foreground Pattern: " + str(override_data['surface_foreground_pattern']))
        if override_data.get('surface_foreground_pattern_color') and override_data['surface_foreground_pattern_color'] != 'Default':
            pattern_parts.append("Foreground Color: " + str(override_data['surface_foreground_pattern_color']))
        if override_data.get('surface_background_pattern') and override_data['surface_background_pattern'] != 'Default':
            pattern_parts.append("Background Pattern: " + str(override_data['surface_background_pattern']))
        if override_data.get('surface_background_pattern_color') and override_data['surface_background_pattern_color'] != 'Default':
            pattern_parts.append("Background Color: " + str(override_data['surface_background_pattern_color']))
        if pattern_parts:
            summary_parts.append("Surface Patterns: " + ", ".join(pattern_parts))
        
        # Transparency
        if override_data.get('transparency'):
            summary_parts.append("Transparency: " + str(override_data['transparency']))
        
        # Cut - Lines
        cut_line_parts = []
        if override_data.get('cut_line_weight'):
            weight_value = override_data['cut_line_weight']
            # Line weight -1 means using object style definition (no local override)
            if weight_value == -1:
                # cut_line_parts.append("Line Weight: Default")
                pass
            else:
                cut_line_parts.append("Line Weight: " + str(weight_value))
        if override_data.get('cut_line_color') and override_data['cut_line_color'] != 'Default':
            cut_line_parts.append("Line Color: " + str(override_data['cut_line_color']))
        if override_data.get('cut_line_pattern') and override_data['cut_line_pattern'] != 'Default':
            cut_line_parts.append("Line Pattern: " + str(override_data['cut_line_pattern']))
        if cut_line_parts:
            summary_parts.append("Cut Lines: " + ", ".join(cut_line_parts))
        
        # Cut - Patterns (Cut Foreground and Background)
        cut_pattern_parts = []
        if override_data.get('cut_foreground_pattern') and override_data['cut_foreground_pattern'] != 'Default':
            cut_pattern_parts.append("Foreground Pattern: " + str(override_data['cut_foreground_pattern']))
        if override_data.get('cut_foreground_pattern_color') and override_data['cut_foreground_pattern_color'] != 'Default':
            cut_pattern_parts.append("Foreground Color: " + str(override_data['cut_foreground_pattern_color']))
        if override_data.get('cut_background_pattern') and override_data['cut_background_pattern'] != 'Default':
            cut_pattern_parts.append("Background Pattern: " + str(override_data['cut_background_pattern']))
        if override_data.get('cut_background_pattern_color') and override_data['cut_background_pattern_color'] != 'Default':
            cut_pattern_parts.append("Background Color: " + str(override_data['cut_background_pattern_color']))
        if cut_pattern_parts:
            summary_parts.append("Cut Patterns: " + ", ".join(cut_pattern_parts))
        
        # Halftone
        if override_data.get('halftone'):
            summary_parts.append("Halftone")
        
        # Detail Level
        if override_data.get('detail_level') and override_data['detail_level'] != 'By View':
            summary_parts.append("Detail: " + str(override_data['detail_level']))
        
        return "\n".join(summary_parts) if summary_parts else "Default Settings"
    
    def _process_summary_for_bold_headers(self, summary):
        """
        Process summary text to make section headers bold using CSS classes and add color blocks.
        
        Args:
            summary: The summary text
            
        Returns:
            str: Processed summary with bold headers and color blocks
        """
        if not summary:
            return summary
            
        # Define section headers that should be bold
        section_headers = [
            "Projection Lines:",
            "Surface Patterns:",
            "Cut Lines:",
            "Cut Patterns:",
            "Visibility:",
            "Transparency:",
            "Halftone",
            "Detail:"
        ]
        
        processed_summary = summary
        
        # Replace each section header with a bold version
        for header in section_headers:
            if header in processed_summary:
                # Use CSS class for bold styling
                bold_header = '<span class="bold-text">{}<br></span>'.format(header)
                processed_summary = processed_summary.replace(header, bold_header)
        
        # Add color blocks next to RGB color values
        import re
        
        # Pattern to match RGB color values like "RGB(255, 0, 0)" or "RGB(255,255,255)"
        rgb_pattern = r'RGB\((\d+),\s*(\d+),\s*(\d+)\)'
        
        def replace_with_color_block(match):
            r, g, b = int(match.group(1)), int(match.group(2)), int(match.group(3))
            color_hex = '#{:02x}{:02x}{:02x}'.format(r, g, b)
            return '{} <span style="display: inline-block; width: 50px; height: 15px; background-color: {}; border: 2px solid #666; border-radius: 2px; vertical-align: middle; margin-left: 5px;" title="{}"></span>'.format(match.group(0), color_hex, color_hex)
        
        processed_summary = re.sub(rgb_pattern, replace_with_color_block, processed_summary)
        
        return processed_summary
    
    def _create_difference_summary(self, override_data, all_template_values):
        """
        Create a summary string showing what properties have inconsistencies across templates.
        
        Logic: For each property, check if there are ANY inconsistencies across templates.
        If a property has inconsistencies, include it in the summary.
        
        Args:
            override_data: Dictionary containing override details for this template
            all_template_values: Dictionary containing override data for all templates
            
        Returns:
            str: Summary string showing properties with inconsistencies
        """
        try:
            if not isinstance(override_data, dict):
                return "Invalid Data"
            
            # Properties to check for differences
            properties_to_check = [
                'projection_line_weight', 'projection_line_color', 'projection_line_pattern',
                'projection_fill_pattern', 'projection_fill_color', 'transparency',
                'cut_line_weight', 'cut_line_color', 'cut_line_pattern',
                'cut_fill_pattern', 'cut_fill_color', 'halftone', 'detail_level'
            ]
            
            # Find properties that have inconsistencies across templates
            inconsistent_properties = []
            
            for prop in properties_to_check:
                # Get all values for this property across all templates
                all_values = []
                for template_name, template_data in all_template_values.items():
                    if isinstance(template_data, dict):
                        value = template_data.get(prop)
                        if value is not None:
                            all_values.append(value)
                
                # Check if there are inconsistencies (more than one unique value)
                if len(all_values) > 1:
                    unique_values = set(all_values)
                    if len(unique_values) > 1:
                        # This property has inconsistencies
                        inconsistent_properties.append(prop)
            
            # Build summary of inconsistent properties
            if not inconsistent_properties:
                return "No Override"
            
            summary_parts = []
            for prop in inconsistent_properties:
                if prop == 'projection_line_weight':
                    summary_parts.append("Line Weight")
                elif prop == 'projection_line_color':
                    summary_parts.append("Line Color")
                elif prop == 'projection_line_pattern':
                    summary_parts.append("Line Pattern")
                elif prop == 'projection_fill_pattern':
                    summary_parts.append("Fill Pattern")
                elif prop == 'projection_fill_color':
                    summary_parts.append("Fill Color")
                elif prop == 'transparency':
                    summary_parts.append("Transparency")
                elif prop == 'cut_line_weight':
                    summary_parts.append("Cut Line Weight")
                elif prop == 'cut_line_color':
                    summary_parts.append("Cut Line Color")
                elif prop == 'cut_line_pattern':
                    summary_parts.append("Cut Line Pattern")
                elif prop == 'cut_fill_pattern':
                    summary_parts.append("Cut Fill Pattern")
                elif prop == 'cut_fill_color':
                    summary_parts.append("Cut Fill Color")
                elif prop == 'halftone':
                    summary_parts.append("Halftone")
                elif prop == 'detail_level':
                    summary_parts.append("Detail Level")
            
            return "; ".join(summary_parts)
            
        except Exception as e:
            return "Error: {}".format(str(e))
    
    def _generate_category_visibility_section(self, differences):
        """
        Generate the category visibility comparison section.
        
        Args:
            differences: Dictionary of category visibility differences
            
        Returns:
            str: HTML for category visibility section
        """
        return self._generate_simple_comparison_section('category_visibility', 'Category Visibility', differences)
    
    def _generate_workset_visibility_section(self, differences):
        """
        Generate the workset visibility comparison section.
        
        Args:
            differences: Dictionary of workset visibility differences
            
        Returns:
            str: HTML for workset visibility section
        """
        return self._generate_simple_comparison_section('workset_visibility', 'Workset Visibility', differences)
    
    def _generate_view_parameters_section(self, differences):
        """
        Generate the view parameters differences section with values.
        
        Args:
            differences: Dictionary of view parameter differences with values
            
        Returns:
            str: HTML for view parameters section
        """
        try:
            # Ensure template_names is a valid list
            if not isinstance(self.template_names, list):
                self.template_names = []
            
            html = """
        <div class="section">
            <h2 class="toggle" onclick="toggleSection('view_parameters')">View Parameters (Click to expand)</h2>
            <div id="view_parameters" class="collapsible">
                <table>
                    <tr>
                        <th>Parameter</th>
"""
            
            for template_name in self.template_names:
                try:
                    if template_name:
                        html += "                        <th>{}</th>\n".format(str(template_name))
                    else:
                        html += "                        <th>Unknown</th>\n"
                except Exception as e:
                    ERROR_HANDLE.print_note("Error processing template name in view parameters: {}".format(str(e)))
                    html += "                        <th>Error</th>\n"
            
            html += "                    </tr>\n"
            
            for param_name, values in differences.items():
                try:
                    html += "                    <tr><td class='template-col'>{}</td>".format(str(param_name))
                    for template_name in self.template_names:
                        try:
                            value = values.get(template_name, 'N/A')
                            if value == "Not Controlled":
                                html += "<td style='background-color: #dc3545; color: white; font-weight: bold; text-align: center;'>{} <span style='color: #FFD700;'>DANGEROUS</span></td>".format(str(value))
                            elif value == "Not Present":
                                html += "<td style='background-color: #e9ecef; color: #495057; font-style: italic; text-align: center;'>{}</td>".format(str(value))
                            else:
                                # Show the actual parameter value
                                html += "<td style='background-color: #d1ecf1; color: #0c5460; font-weight: bold; text-align: center;'>{}</td>".format(str(value))
                        except Exception as e:
                            ERROR_HANDLE.print_note("Error processing parameter value for template {}: {}".format(template_name, str(e)))
                            html += "<td class='different'>Error</td>"
                    html += "</tr>\n"
                except Exception as e:
                    ERROR_HANDLE.print_note("Error processing parameter {}: {}".format(param_name, str(e)))
                    continue
            
            html += "                </table></div></div>\n"
            return html
            
        except Exception as e:
            ERROR_HANDLE.print_note("Error generating view parameters section: {}".format(str(e)))
            return "<h2 style='color: red;'>Error Generating View Parameters Section</h2>\n"
    
    def _generate_uncontrolled_parameters_section(self, differences):
        """
        Generate the uncontrolled parameters section (DANGEROUS).
        
        Args:
            differences: Dictionary of uncontrolled parameter differences
            
        Returns:
            str: HTML for uncontrolled parameters section
        """
        html = """
        <div class="section">
            <h2 class="toggle" onclick="toggleSection('uncontrolled_parameters')" style="color: red;">Uncontrolled Parameters - DANGEROUS (Click to expand)</h2>
            <div id="uncontrolled_parameters" class="collapsible">
                <div style="background-color: #8B0000; border: 2px solid #ff0000; padding: 15px; margin-bottom: 15px; border-radius: 5px; color: #ffffff;">
                    <strong style="color: #FFD700;">WARNING:</strong> Uncontrolled parameters are NOT marked as included when the view is used as a template. 
                    This can cause inconsistent behavior across views using the same template. 
                    These parameters should be reviewed and controlled if needed for consistency.
                </div>
                <table>
                    <tr>
                        <th>Parameter</th>
"""
        
        for template_name in self.template_names:
            html += "                        <th>{}</th>\n".format(template_name)
        
        html += "                    </tr>\n"
        
        for param, values in differences.items():
            html += "                    <tr><td class='template-col'>{}</td>".format(param)
            for template_name in self.template_names:
                value = values.get(template_name, False)
                if value == "UNCONTROLLED":
                    html += "<td style='background-color: #dc3545; color: white; font-weight: bold; text-align: center;'>UNCONTROLLED</td>"
                elif value:
                    html += "<td style='background-color: #f8d7da; color: #721c24; font-weight: bold; text-align: center;'>Uncontrolled</td>"
                else:
                    html += "<td style='background-color: #d4edda; color: #155724; font-weight: bold; text-align: center;'>Controlled</td>"
            html += "</tr>\n"
        
        html += "                </table></div></div>\n"
        return html
    
    def _generate_filters_section(self, differences):
        """
        Generate the filters comparison section.
        
        Args:
            differences: Dictionary of filter differences
            
        Returns:
            str: HTML for filters section
        """
        html = """
        <div class="section">
            <h2 class="toggle" onclick="toggleSection('filters')">Filters (Click to expand)</h2>
            <div id="filters" class="collapsible">
                <table>
                    <tr>
                        <th>Filter</th>
"""
        
        for template_name in self.template_names:
            html += "                        <th>{}</th>\n".format(template_name)
        
        html += "                    </tr>\n"
        
        for filter_name, values in differences.items():
            html += "                    <tr><td class='template-col'>{}</td>".format(filter_name)
            for template_name in self.template_names:
                value = values.get(template_name, 'N/A')
                if value == "UNCONTROLLED":
                    html += "<td style='background-color: #FF8C00; color: white; font-weight: bold;'>UNCONTROLLED</td>"
                elif value:
                    summary = self._create_override_summary(value)
                    html += "<td class='different' style='white-space: pre-line;'>{}</td>".format(summary)
                else:
                    html += "<td class='same'>Not Applied</td>"
            html += "</tr>\n"
        
        html += "                </table></div></div>\n"
        return html
    
    def _generate_template_usage_section(self, comparison_data=None):
        """
        Generate the template usage section showing which views use each template.
        
        Args:
            comparison_data: Dictionary containing template data with usage information
            
        Returns:
            str: HTML for template usage section
        """
        try:
            html = """
        <div class="section">
            <h2 class="toggle" onclick="toggleSection('template_usage')">Template Usage - Views Using Each Template (Click to expand)</h2>
            <div id="template_usage" class="collapsible">
"""
            
            if comparison_data and isinstance(comparison_data, dict):
                # Generate usage data for each template
                for template_name in self.template_names:
                    template_data = comparison_data.get(template_name, {})
                    usage_data = template_data.get('template_usage', {})
                    
                    html += """
                <div style="margin-bottom: 30px; padding: 20px; background: #333333; border-radius: 8px; border: 1px solid #404040;">
                    <h3 style="color: #ffffff; margin-top: 0; margin-bottom: 15px; font-size: 1.3em;">{}</h3>
""".format(template_name)
                    
                    views = usage_data.get('views', [])
                    total_count = usage_data.get('total_count', 0)
                    
                    if views:
                        html += """
                    <p style="color: #cccccc; margin-bottom: 15px;"><strong>Total Views Using This Template: {}</strong></p>
                    <table style="width: 100%; margin-bottom: 15px;">
                        <tr>
                            <th style="background: #404040; color: #ffffff; padding: 10px; text-align: left;">View Name</th>
                            <th style="background: #404040; color: #ffffff; padding: 10px; text-align: left;">View Type</th>
                            <th style="background: #404040; color: #ffffff; padding: 10px; text-align: left;">View ID</th>
                        </tr>
""".format(total_count)
                        
                        for view in views:
                            html += """
                        <tr>
                            <td style="padding: 8px; border-bottom: 1px solid #404040; color: #ffffff;">{}</td>
                            <td style="padding: 8px; border-bottom: 1px solid #404040; color: #cccccc;">{}</td>
                            <td style="padding: 8px; border-bottom: 1px solid #404040; color: #999999; font-family: 'JetBrains Mono', monospace; font-size: 0.9em;">{}</td>
                        </tr>
""".format(view.get('name', 'Unknown'), view.get('type', 'Unknown'), view.get('id', 'Unknown'))
                        
                        html += "                    </table>"
                    else:
                        html += """
                    <p style="color: #999999; font-style: italic;">No views are currently using this template.</p>
"""
                    
                    html += "                </div>"
            else:
                html += """
                <p style="color: #cccccc; font-style: italic;">Template usage data is not available.</p>
                <p style="color: #999999; font-size: 0.9em;">This section shows which views are currently using each template, sorted alphabetically by view name.</p>
"""
            
            # Add footnote about upcoming feature
            html += """
                <div style="margin-top: 20px; padding: 15px; background: #2a2a2a; border-radius: 8px; border: 1px solid #404040;">
                    <p style="color: #999999; font-size: 0.9em; margin: 0;">
                        <strong>Note:</strong> Template usage detection is currently in development. 
                        This feature will show which views are using each template in future updates.
                    </p>
                </div>
            </div></div>\n"""
            return html
            
        except Exception as e:
            return """
        <div class="section">
            <h2 style="color: red;">Error Generating Template Usage Section</h2>
            <div class="error">
                <p>An error occurred while generating template usage section:</p>
                <p><strong>{}</strong></p>
            </div>
        </div>
""".format(str(e))
    
    def _generate_simple_comparison_section(self, section_id, section_title, differences):
        """
        Generate a simple comparison section for basic values.
        
        Args:
            section_id: ID for the HTML section
            section_title: Title for the section
            differences: Dictionary of differences
            
        Returns:
            str: HTML for the section
        """
        try:
            # Ensure template_names is a valid list
            if not isinstance(self.template_names, list):
                self.template_names = []
            
            html = """
        <div class="section">
            <h2 class="toggle" onclick="toggleSection('{}')">{} (Click to expand)</h2>
            <div id="{}" class="collapsible">
                <table>
                    <tr>
                        <th>Item</th>
""".format(section_id, section_title, section_id)
            
            for template_name in self.template_names:
                try:
                    if template_name:
                        html += "                        <th>{}</th>\n".format(str(template_name))
                    else:
                        html += "                        <th>Unknown</th>\n"
                except Exception as e:
                    ERROR_HANDLE.print_note("Error processing template name in simple comparison: {}".format(str(e)))
                    html += "                        <th>Error</th>\n"
            
            html += "                    </tr>\n"
            
            for item, values in differences.items():
                try:
                    html += "                    <tr><td class='template-col'>{}</td>".format(str(item))
                    for template_name in self.template_names:
                        try:
                            value = values.get(template_name, 'N/A')
                            if value == "UNCONTROLLED":
                                html += "<td style='background-color: #dc3545; color: white; font-weight: bold;'>UNCONTROLLED <span style='color: #FFD700;'>DANGEROUS</span></td>"
                            elif value == "Visible" or value == "On":
                                html += "<td class='visible-cell'>{}</td>".format(str(value))
                            elif value == "Hidden":
                                html += "<td class='hidden-cell'>{}</td>".format(str(value))
                            else:
                                html += "<td class='different'>{}</td>".format(str(value))
                        except Exception as e:
                            ERROR_HANDLE.print_note("Error processing value for template {}: {}".format(template_name, str(e)))
                            html += "<td class='different'>Error</td>"
                    html += "</tr>\n"
                except Exception as e:
                    ERROR_HANDLE.print_note("Error processing item {}: {}".format(item, str(e)))
                    continue
            
            html += "                </table></div></div>\n"
            return html
                
        except Exception as e:
            ERROR_HANDLE.print_note("Error generating simple comparison section: {}".format(str(e)))
            return """
        <div class="section">
            <h2 class="toggle" onclick="toggleSection('{}')">{} (Click to expand)</h2>
            <div id="{}" class="collapsible">
                <div class="error">
                    <p>Error generating {} section: {}</p>
                </div>
            </div>
        </div>
""".format(section_id, section_title, section_id, section_title, str(e))
    
    def _generate_boolean_comparison_section(self, section_id, section_title, differences, true_text, false_text):
        """
        Generate a boolean comparison section.
        
        Args:
            section_id: ID for the HTML section
            section_title: Title for the section
            differences: Dictionary of differences
            true_text: Text to display for True values
            false_text: Text to display for False values
            
        Returns:
            str: HTML for the section
        """
        try:
            # Ensure template_names is a valid list
            if not isinstance(self.template_names, list):
                self.template_names = []
            
            html = """
        <div class="section">
            <h2 class="toggle" onclick="toggleSection('{}')">{} (Click to expand)</h2>
            <div id="{}" class="collapsible">
                <table>
                    <tr>
                        <th>Parameter</th>
""".format(section_id, section_title, section_id)
            
            for template_name in self.template_names:
                try:
                    if template_name:
                        html += "                        <th>{}</th>\n".format(str(template_name))
                    else:
                        html += "                        <th>Unknown</th>\n"
                except Exception as e:
                    ERROR_HANDLE.print_note("Error processing template name in boolean comparison: {}".format(str(e)))
                    html += "                        <th>Error</th>\n"
            
            html += "                    </tr>\n"
            
            for item, values in differences.items():
                try:
                    html += "                    <tr><td class='template-col'>{}</td>".format(str(item))
                    for template_name in self.template_names:
                        try:
                            value = values.get(template_name, False)
                            if value == "UNCONTROLLED":
                                html += "<td style='background-color: #dc3545; color: white; font-weight: bold;'>UNCONTROLLED <span style='color: #FFD700;'>DANGEROUS</span></td>"
                            elif value:
                                html += "<td class='same'>{}</td>".format(str(true_text))
                            else:
                                html += "<td class='different'>{}</td>".format(str(false_text))
                        except Exception as e:
                            ERROR_HANDLE.print_note("Error processing boolean value for template {}: {}".format(template_name, str(e)))
                            html += "<td class='different'>Error</td>"
                    html += "</tr>\n"
                except Exception as e:
                    ERROR_HANDLE.print_note("Error processing boolean item {}: {}".format(item, str(e)))
                    continue
            
            html += "                </table></div></div>\n"
            return html
            
        except Exception as e:
            ERROR_HANDLE.print_note("Error generating boolean comparison section: {}".format(str(e)))
            return """
        <div class="section">
            <h2 class="toggle" onclick="toggleSection('{}')">{} (Click to expand)</h2>
            <div id="{}" class="collapsible">
                <div class="error">
                    <p>Error generating {} section: {}</p>
                </div>
            </div>
        </div>
""".format(section_id, section_title, section_id, section_title, str(e))
    
    def _generate_html_footer(self):
        """
        Generate the HTML footer.
        
        Returns:
            str: HTML footer
        """
        return """
    </div>
</body>
</html>
"""
    
    def save_report_to_file(self, html_content):
        """
        Save the HTML report to the DUMP folder and return the filepath.
        
        Args:
            html_content: The HTML content to save
            
        Returns:
            str: Filepath of the saved HTML file
        """
        
        # Create filename without timestamp
        filename = "ViewTemplate_Comparison.html"
        filepath = os.path.join(FOLDER.DUMP_FOLDER, filename)
        
        # Write HTML content to file with proper IronPython 2.7 Unicode handling
        try:
            # IronPython 2.7 approach: encode string to UTF-8 and write as binary
            try:
                # Check if unicode type exists (IronPython 2.7)
                if 'unicode' in dir(__builtins__) and isinstance(html_content, globals()['unicode']):
                    # If it's already unicode, encode it
                    html_bytes = html_content.encode('utf-8')
                else:
                    # If it's a regular string, decode then re-encode to ensure UTF-8
                    html_bytes = html_content.decode('utf-8').encode('utf-8')
            except NameError:
                # unicode type doesn't exist (Python 3)
                html_bytes = html_content.encode('utf-8')
            
            with open(filepath, 'wb') as f:
                f.write(html_bytes)
                
        except (UnicodeDecodeError, UnicodeEncodeError, NameError):
            # Fallback for cases where unicode type doesn't exist (Python 3) or encoding fails
            try:
                # Try simple UTF-8 encoding
                html_content_encoded = html_content.encode('utf-8')
                with open(filepath, 'wb') as f:
                    f.write(html_content_encoded)
            except (UnicodeEncodeError, AttributeError):
                # Last resort: ASCII fallback
                ERROR_HANDLE.print_note("Warning: Falling back to ASCII encoding for compatibility")
                html_content_safe = str(html_content).encode('ascii', 'replace').decode('ascii')
                with open(filepath, 'w') as f:
                    f.write(html_content_safe)
        
        ERROR_HANDLE.print_note("HTML report saved to: {}".format(filepath))
        
        return filepath
    
    def _clean_problematic_unicode(self, data):
        """
        Recursively clean only problematic Unicode surrogates that cause encoding issues.
        Preserves normal Unicode characters.
        
        Args:
            data: Data structure (dict, list, string, etc.) to clean
            
        Returns:
            Cleaned data structure with problematic surrogates replaced
        """
        if isinstance(data, dict):
            cleaned_dict = {}
            for key, value in data.items():
                # Clean both key and value
                clean_key = self._clean_problematic_string(str(key)) if key is not None else str(key)
                clean_value = self._clean_problematic_unicode(value)
                cleaned_dict[clean_key] = clean_value
            return cleaned_dict
        elif isinstance(data, list):
            return [self._clean_problematic_unicode(item) for item in data]
        elif isinstance(data, str):
            return self._clean_problematic_string(data)
        else:
            return data
    
    def _clean_problematic_string(self, text):
        """
        Clean only problematic Unicode surrogates from a string.
        Preserves normal Unicode characters that are properly encoded.
        
        Args:
            text: String to clean
            
        Returns:
            String with problematic surrogates replaced
        """
        if text is None:
            return None
        try:
            # Convert to string first if needed
            text_str = str(text)
            # Only replace specific problematic Unicode surrogates
            # These are the high/low surrogate pairs that cause issues in IronPython
            cleaned = text_str.replace(u'\uD83D\uDE00', '[smile]')  # 😀
            cleaned = cleaned.replace(u'\uD83D\uDE01', '[grin]')   # 😁
            cleaned = cleaned.replace(u'\uD83D\uDC41', '[eye]')    # 👁
            cleaned = cleaned.replace(u'\uD83D\uDEAB', '[no]')     # 🚫
            # Replace any remaining high surrogates
            cleaned = cleaned.replace(u'\uD83D', '[emoji]')
            return cleaned
        except (UnicodeEncodeError, UnicodeDecodeError):
            # If all else fails, return original text - let UTF-8 encoding handle it
            return text


if __name__ == "__main__":
    pass