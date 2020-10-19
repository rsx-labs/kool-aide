@echo off
"C:\Program Files (x86)\kool-aide\kool-aide" generate-report -r resource-planner --format excel --output "e:\\AIDE Generated Reports\\Planner\\Retail Services DEV Resource Planner for [LM][Y].xlsx" --params {\"departments\":[1],\"divisions\":[1]} --autorun
@echo off
"C:\Program Files (x86)\kool-aide\kool-aide" generate-report -r resource-planner --format excel --output "e:\\AIDE Generated Reports\\Planner\\Retail Services QA Resource Planner for [LM][Y].xlsx" --params {\"departments\":[1],\"divisions\":[2]} --autorun