import os
import json
import subprocess
import statistics
import argparse
from datetime import datetime
from docx import Document
from docx.shared import RGBColor

# Define paths to ESLint and complexity-report within the CODE_ANALYSIS project
ESLINT_PATH = os.path.join('node_modules', '.bin', 'eslint')
COMPLEXITY_REPORT_PATH = os.path.join('node_modules', '.bin', 'cr.cmd')  # Adjust for Unix/Mac if needed

def analyze_code(file_path):
    """Runs ESLint and complexity-report analysis on a single file and captures the results."""
    
    # Run ESLint Analysis
    try:
        eslint_output = subprocess.run(
            [ESLINT_PATH, file_path, '-f', 'json'],
            capture_output=True, text=True, check=True, shell=True
        )
        eslint_results = json.loads(eslint_output.stdout)
        eslint_issues = len(eslint_results[0]['messages']) if eslint_results else 0
    except subprocess.CalledProcessError as e:
        print(f"Error running ESLint on {file_path}: {e}")
        eslint_issues = 0
    except FileNotFoundError as fnf_error:
        print(f"ESLint not found: {fnf_error}")
        eslint_issues = 0

    # Run Complexity-Report Analysis
    try:
        if not os.path.exists(COMPLEXITY_REPORT_PATH):
            raise FileNotFoundError(f"The complexity-report executable was not found at: {COMPLEXITY_REPORT_PATH}")

        complexity_output = subprocess.run(
            [COMPLEXITY_REPORT_PATH, '--output', 'json', file_path],
            capture_output=True, text=True, check=True, shell=True
        )

        if not complexity_output.stdout.strip():
            raise ValueError(f"No output produced by complexity-report for {file_path}. It might be caused by unsupported syntax or an internal error.")

        complexity_data = json.loads(complexity_output.stdout)
        
        if complexity_data and 'functions' in complexity_data:
            functions = complexity_data['functions']
            cc_scores = [func['cyclomatic'] for func in functions]
            avg_cc = sum(cc_scores) / len(cc_scores) if cc_scores else 0
            max_cc = max(cc_scores) if cc_scores else 0
            
            maintainability_index = complexity_data.get('maintainability', 100)
            halstead_metrics = complexity_data.get('aggregate', {}).get('halstead', {})
            volume = halstead_metrics.get('volume', 0)
            difficulty = halstead_metrics.get('difficulty', 0)
            effort = halstead_metrics.get('effort', 0)
            vocabulary = halstead_metrics.get('vocabulary', 0)
            length = halstead_metrics.get('length', 0)
            total_operators = halstead_metrics.get('operators', {}).get('total', 0)
            total_operands = halstead_metrics.get('operands', {}).get('total', 0)
        else:
            avg_cc = max_cc = volume = difficulty = effort = vocabulary = length = 0
            total_operators = total_operands = maintainability_index = 100
            
    except subprocess.CalledProcessError as e:
        print(f"Error running complexity-report on {file_path}:\n{e.stderr}")
        log_error(file_path, e.stderr)
        avg_cc = max_cc = volume = difficulty = effort = vocabulary = length = 0
        total_operators = total_operands = maintainability_index = 100
    except (FileNotFoundError, ValueError, json.JSONDecodeError) as e:
        print(f"Error processing file {file_path}: {e}")
        log_error(file_path, str(e))
        avg_cc = max_cc = volume = difficulty = effort = vocabulary = length = 0
        total_operators = total_operands = maintainability_index = 100

    return {
        "file_path": file_path,
        "ESLint Issues": eslint_issues,
        "Cyclomatic Complexity (avg)": avg_cc,
        "Cyclomatic Complexity (max)": max_cc,
        "Halstead Vocabulary": vocabulary,
        "Halstead Length": length,
        "Halstead Volume": volume,
        "Halstead Effort": effort,
        "Halstead Difficulty": difficulty,
        "Total Operators": total_operators,
        "Total Operands": total_operands,
        "Maintainability Index": maintainability_index
    }

def log_error(file_path, error_message):
    """Logs errors to a file for further inspection."""
    with open("complexity_report_errors.log", "a", encoding="utf-8") as log_file:
        log_file.write(f"Error analyzing {file_path}:\n{error_message}\n{'-' * 80}\n")

def analyze_project_folder(folder_path):
    """Analyzes all JavaScript/TypeScript files in a given project folder."""
    all_metrics = []
    
    for root, dirs, files in os.walk(folder_path):
        # Ignore 'node_modules' and other folders you don't want to analyze
        dirs[:] = [d for d in dirs if d not in ['node_modules']]
        
        for file in files:
            if file.endswith(('.js', '.jsx', '.ts', '.tsx')):
                file_path = os.path.join(root, file)
                print(f"Analyzing: {file_path}")
                metrics = analyze_code(file_path)
                all_metrics.append(metrics)
    
    if not all_metrics:
        print("No JavaScript/TypeScript files found in the specified folder.")
        
    return all_metrics

def aggregate_metrics(all_metrics):
    """Aggregates metrics for the entire project."""
    if not all_metrics:
        print("No metrics to aggregate - no JavaScript/TypeScript files were analyzed.")
        return {}
    
    aggregated = {
        "Total Files": len(all_metrics),
        "Total ESLint Issues": sum([m["ESLint Issues"] for m in all_metrics]),
        "Average ESLint Issues per File": statistics.mean([m["ESLint Issues"] for m in all_metrics]),
        "Average Cyclomatic Complexity": statistics.mean([m["Cyclomatic Complexity (avg)"] for m in all_metrics]),
        "Maximum Cyclomatic Complexity": max([m["Cyclomatic Complexity (max)"] for m in all_metrics]),
        "Average Halstead Volume": statistics.mean([m["Halstead Volume"] for m in all_metrics]),
        "Average Halstead Effort": statistics.mean([m["Halstead Effort"] for m in all_metrics]),
        "Average Halstead Difficulty": statistics.mean([m["Halstead Difficulty"] for m in all_metrics]),
        "Average Maintainability Index": statistics.mean([m["Maintainability Index"] for m in all_metrics])
    }
    return aggregated

def generate_report(all_file_metrics, aggregated_metrics, output_path, format_type):
    """Generate either an HTML or Word report based on the format_type parameter."""
    if format_type == 'html':
        generate_html_report(all_file_metrics, aggregated_metrics, output_path)
    else:
        generate_word_report(all_file_metrics, aggregated_metrics, output_path)

def generate_html_report(all_file_metrics, aggregated_metrics, output_path):
    """Generates an HTML report from the analysis results."""
    html_content = f"""
    <html>
    <head>
        <title>React Code Analysis Report</title>
        <style>
            body {{ font-family: Arial, sans-serif; }}
            .container {{ padding: 20px; }}
            table {{ border-collapse: collapse; width: 100%; margin-bottom: 20px; }}
            th, td {{ border: 1px solid #ddd; padding: 8px; }}
            th {{ background-color: #f2f2f2; }}
            .good {{ background-color: #d4edda; color: #155724; }}
            .moderate {{ background-color: #fff3cd; color: #856404; }}
            .poor {{ background-color: #f8d7da; color: #721c24; }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>React Code Analysis Report</h1>
            <h2>Aggregated Project Metrics</h2>
            <table>
                <tr><th>Metric</th><th>Value</th></tr>
    """
    for metric, value in aggregated_metrics.items():
        html_content += f"<tr><td>{metric}</td><td>{value}</td></tr>"

    html_content += "</table><h2>Individual File Metrics</h2>"
    
    for file_metrics in all_file_metrics:
        html_content += f"""
        <h3>File: {file_metrics['file_path']}</h3>
        <table>
            <tr><th>Metric</th><th>Value</th></tr>
        """
        for metric, value in file_metrics.items():
            if metric != "file_path":
                html_content += f"<tr><td>{metric}</td><td>{value}</td></tr>"
        
        html_content += "</table>"
    
    html_content += "</div></body></html>"

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print(f"HTML report generated: {output_path}")

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Generate React code analysis reports in HTML or Word format")
    parser.add_argument('folder_path', type=str, help="Path to the React project src folder")
    parser.add_argument('--format', choices=['html', 'word'], default='html', help="The report format to generate (html or word)")
    parser.add_argument('--output', type=str, required=True, help="The output path for the report")
    args = parser.parse_args()
    
    folder_path = args.folder_path
    report_format = args.format
    output_path = args.output

    all_file_metrics = analyze_project_folder(folder_path)
    aggregated_metrics = aggregate_metrics(all_file_metrics)

    generate_report(all_file_metrics, aggregated_metrics, output_path, report_format)
