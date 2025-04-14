Awesome, Josh. Since Tachyon AI uses RAG (Retrieval-Augmented Generation), this is a perfect scenario for leveraging both your exported data (changes + suppression matrix) and strategic prompting.

---

## Recommended Setup for Your Use Case
→ Upload both documents:
- `changes.csv` → Weekly exported ServiceNow changes
- `suppression_matrix.xlsx` → Suppression rules per application

---

## Suggested System Prompt (Sets AI Behavior & Role)
```markdown
You are an expert IT Change Management AI assistant helping a Platform Engineer analyze weekly change records from ServiceNow against a suppression matrix.

Your responsibilities:
1. Identify which changes require alert suppression, partner notifications, banners, or special considerations.
2. Match the "Configuration Item" in the changes data with the "Application" in the suppression matrix (ignore case sensitivity and extra spaces).
3. If a Configuration Item is not found in the suppression matrix, indicate "Not Found in Suppression Matrix".
4. Provide clear output tables or bullet lists summarizing required actions.
5. When summarizing results, include Change Number, Configuration Item, Planned Start & End Dates, and required suppression actions.
6. If multiple actions are required (Suppression, Banner, Partner Notification), list them all.
7. Be concise but complete in all responses.
```

---

## Example User Prompts (Per Task)

### 1. Show me all changes that require suppression this week
> "List all changes from the changes document that require suppression based on the suppression matrix. Include suppression details and planned dates."

---

### 2. Which changes require partner notification or banner updates?
> "Identify changes where either 'Notification to Partners' or 'Banner' in the suppression matrix is marked as required. Summarize with Change Number, Configuration Item, and action required."

---

### 3. Are there any configuration items not mapped to suppression rules?
> "List all Configuration Items from the changes document that do not exist in the suppression matrix."

---

### 4. Give me a risk summary of all critical changes needing suppression
> "Show changes with Risk marked as 'High' or Category as 'Infrastructure' that also require suppression."

---

### 5. Prepare a report with all changes and their suppression requirements
> "Create a table of all changes from this week, with columns: Change Number, Configuration Item, Suppression Needed (Yes/No), Suppression Records, Partner Notification, Banner Needed, Outage Impact."

---

## Bonus: Output Format Suggestions
Ask Tachyon to output results like this for easy copy-paste to Excel or Teams:

| Change Number | Config Item | Start Date | End Date | Suppression | Partner Notification | Banner Needed |
|---------------|-------------|------------|----------|-------------|----------------------|----------------|
| CHG1234567    | PAYMENTS API | 2024-04-15 | 2024-04-15 | Yes (Record: Splunk-Payments) | Yes | No |

---
