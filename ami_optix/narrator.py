import json
import os
from openai import OpenAI
from groq import Groq

def generate_internal_summary(analysis_json):
    """
    Generates a simple, fact-based text summary of the analysis results
    without using an external LLM.
    """
    summary_parts = []

    abs_best = analysis_json.get("scenario_absolute_best")
    if abs_best:
        waami = abs_best['waami'] * 100
        summary_parts.append(f"The analysis found an optimal scenario with a WAAMI of {waami:.4f}%.")
    else:
        summary_parts.append("No optimal scenario was found.")

    notes = analysis_json.get("analysis_notes", [])
    if notes:
        summary_parts.append("\nAnalysis Notes:")
        for note in notes:
            summary_parts.append(f"- {note}")

    flagged_reports = [r['details'] for r in analysis_json.get('compliance_report', []) if r['status'] == 'FLAGGED']
    if flagged_reports:
        summary_parts.append("\nCompliance Issues Found:")
        for report in flagged_reports:
            summary_parts.append(f"- {report}")

    return "\n".join(summary_parts)

def _format_scenario_summary(scenario_name, scenario_data):
    """Creates a formatted string summary for a single scenario for the LLM prompt."""
    if not scenario_data:
        return ""

    waami = scenario_data['waami'] * 100
    bands = ", ".join(map(str, scenario_data['bands']))

    summary = f"### {scenario_name}:\n"
    summary += f"- Final WAAMI: {waami}\n"
    summary += f"- Bands Used: {bands}\n"

    assignment_counts = {}
    for unit in scenario_data['assignments']:
        ami = unit['assigned_ami']
        assignment_counts[ami] = assignment_counts.get(ami, 0) + 1

    breakdown_parts = []
    for ami, count in sorted(assignment_counts.items()):
        band_percent = int(ami * 100)
        breakdown_parts.append(f"{count} units at {band_percent} AMI")

    summary += f"- Unit Assignment Summary: {', '.join(breakdown_parts)}.\n"

    return summary

def _build_prompt(analysis_json):
    """Constructs a detailed, fact-based prompt for the LLM."""
    prompt = "You are an expert real estate financial analyst providing a strategic summary for a client. Your tone is professional, insightful, and clear. Do not repeat the raw numbers verbatim; instead, interpret their meaning.\n\n"
    prompt += "Analyze the following affordable housing scenarios that have been mathematically optimized for a project. Provide a two-paragraph strategic analysis comparing the scenarios. The first paragraph should analyze the 'Absolute Best' scenario against its main 'Alternative'. The second paragraph should analyze the 'Client Oriented' and 'Best 2-Band' scenarios, explaining the strategic trade-offs they offer.\n\n"
    prompt += "--- DATA FOR ANALYSIS ---\n\n"
    prompt += _format_scenario_summary("Scenario 1: Absolute Best", analysis_json.get("scenario_absolute_best"))
    prompt += _format_scenario_summary("Scenario 2: Alternative", analysis_json.get("scenario_alternative"))
    prompt += _format_scenario_summary("Scenario 3: Client Oriented", analysis_json.get("scenario_client_oriented"))
    prompt += _format_scenario_summary("Scenario 4: Best 2-Band", analysis_json.get("scenario_best_2_band"))

    if analysis_json.get("compliance_report"):
        flagged_reports = [r['details'] for r in analysis_json['compliance_report'] if r['status'] == 'FLAGGED']
        if flagged_reports:
            prompt += "\n### Compliance Issues Identified:\n"
            for report in flagged_reports:
                prompt += f"- {report}\n"

    if analysis_json.get("analysis_notes"):
        prompt += "\n### Solver Analysis Notes:\n"
        for note in analysis_json['analysis_notes']:
            prompt += f"- {note}\n"

    prompt += "\n--- END OF DATA ---\n\n"
    prompt += "Begin your two-paragraph strategic analysis now:"
    return prompt

def generate_llm_narrative(analysis_json, provider, model_name):
    """
    Generates a narrative analysis by calling the specified LLM provider.
    """
    prompt = _build_prompt(analysis_json)

    try:
        if provider.lower() == 'openai':
            api_key = os.environ.get("OPENAI_API_KEY")
            if not api_key:
                return "Error: OPENAI_API_KEY environment variable not set."
            client = OpenAI(api_key=api_key)
            response = client.chat.completions.create(
                model=model_name,
                messages=[{"role": "user", "content": prompt}]
            )
            return response.choices[0].message.content

        elif provider.lower() == 'groq':
            api_key = os.environ.get("GROQ_API_KEY")
            if not api_key:
                return "Error: GROQ_API_KEY environment variable not set."
            client = Groq(api_key=api_key)
            response = client.chat.completions.create(
                model=model_name,
                messages=[{"role": "user", "content": prompt}]
            )
            return response.choices[0].message.content

        else:
            return f"Error: Unknown provider '{provider}'. Supported providers are 'openai' and 'groq'."

    except Exception as e:
        return f"Error: An exception occurred while contacting the {provider} API: {e}"
