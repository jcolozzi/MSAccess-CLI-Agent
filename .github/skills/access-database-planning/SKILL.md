# Access Database Planning Skill

Plan, structure, and document Microsoft Access database features using a two-phase approach: Product Requirements Documents (PRD) and task-driven implementation.

## When to Use This Skill

Use this skill when you need to:
- **Plan a new Access feature** from a user request or idea
- **Break down complex requirements** into manageable tasks
- **Document design decisions** before implementation begins
- **Clarify unclear requests** (e.g., missing acceptance criteria, undefined scope, conflicting requirements) before coding starts

## How It Works

This skill uses a streamlined three-step workflow:

**Step 1:** Create a PRD (Product Requirements Document) to clarify what needs to be built and why.  
**Step 2:** Submit PRD for approval before proceeding to implementation planning.  
**Step 3:** Generate task lists with step-by-step implementation guidance only after PRD approval.

| Phase | Goal | Output |
|-------|------|--------|
| **Step 1: PRD** | Clarify **what** needs to be built and **why** | `prd-[feature-name].md` |
| **Step 2: Approval** | Get stakeholder/user sign-off on requirements | PRD review checklist |
| **Step 3: Tasks** | Break approved requirements into implementation steps | `tasks-[feature-name].md` |

---

## Phase 1: Creating a PRD

### When to Create a PRD

Always create a PRD when:
- The user's request is vague or open-ended
- The feature is complex or touches multiple components
- You need stakeholder/user clarity before implementation
- The feature has unclear scope or success criteria
- If user responses are incomplete or contradictory, request additional clarification before proceeding to PRD generation.

### PRD Process

1. **Receive Initial Request:** User provides a brief feature description or request
2. **Ask Clarifying Questions:** Always ask all 4 mandatory questions below for clarity, regardless of request complexity:
   - **Problem/Goal:** What problem does this solve?
   - **Core Functionality:** What can users do with this?
   - **Scope/Boundaries:** What should this NOT do?
   - **Success Criteria:** How do we measure success?
3. **Generate PRD:** Use the PRD Structure (see below) based on user answers
4. **Save:** Store as `prd-[feature-name].md` in the `/tasks` directory

### PRD Structure

The generated PRD must include these **essential sections** (in this order):

**Core Sections (Always Include):**
1. **Introduction/Overview** — What is this feature? What problem does it solve?
2. **Goals** — Specific, measurable objectives
3. **User Stories** — "As a [user type], I want [capability] so that [benefit]"
4. **Functional Requirements** — Numbered list (FR1, FR2, etc.) of must-haves
5. **Non-Goals** — Explicitly state what is OUT of scope
6. **Success Metrics** — How to measure success (e.g., "Reduce support tickets by 20%")
7. **Open Questions** — Remaining ambiguities needing clarification

**Additional Sections (Include if Relevant):**
- Design Considerations — UI/UX, mockups, styling guidelines
- Technical Considerations — Constraints, dependencies, architecture notes

### Clarifying Questions Format

When asking questions, use this format:

```
1. [Question text]
   A. [Option A]
   B. [Option B]
   C. [Option C]
   D. [Option D]

2. [Question text]
   A. [Option A]
   B. [Option B]
```

Users respond with selections like "1A, 2B, 3C" for easy reference.

---

## Phase 2: Generating Task Lists

### When to Create Task Lists

Create a task list when you have:
- An approved PRD
- Clear functional requirements
- A well-defined scope
- Need to guide implementation

### Task List Process

1. **Analyze Requirements:** Review PRD or requirements document
2. **Generate Parent Tasks:** Create 4-6 high-level tasks:
   - **Always include Task 0.0:** "Create feature branch" (unless user specifically opts out)
   - Example parent tasks: "Design Database Schema", "Build Form UI", "Implement Business Logic", "Write Tests", "Documentation"
3. **Present to User:** Show parent tasks and ask "Ready to generate sub-tasks? Respond with 'Go' to proceed"
4. **Generate Sub-Tasks:** Break each parent task into 2-5 actionable sub-tasks
5. **Identify Relevant Files:** List all files that will be created or modified
6. **Save:** Store as `tasks-[feature-name].md` in the `/tasks` directory

### Task List Structure

```markdown
# [Feature Name] - Implementation Tasks

## Relevant Files

- `path/to/file1.accdb` - Brief description of why relevant
- `path/to/module1.bas` - Description of module purpose
- `path/to/form1.frm` - Form code module
- `path/to/test/test_feature.ps1` - PowerShell tests for this feature

### Notes

- Store VBA modules alongside form/report files for version control
- PowerShell tests use the AccessPOSH module for COM automation
- Update this checklist after completing each sub-task (mark `- [ ]` → `- [x]`)

## Tasks

- [ ] 0.0 Create feature branch
  - [ ] 0.1 Create and checkout a new branch (e.g., `git checkout -b feature/[feature-name]`)

- [ ] 1.0 [Parent Task Title]
  - [ ] 1.1 [Sub-task description]
  - [ ] 1.2 [Sub-task description]

- [ ] 2.0 [Parent Task Title]
  - [ ] 2.1 [Sub-task description]
  - [ ] 2.2 [Sub-task description]
```

### Completing Tasks

As you complete each task:
1. Change `- [ ]` to `- [x]` in the markdown
2. Update after completing each sub-task (not just parent tasks)
3. Verify the work passes testing before marking complete

---

## Best Practices

### Key Principles
1. **Clarity First:** Write for a junior developer. Avoid jargon. Focus on "what" and "why," not "how."
2. **Completeness:** Ask essential questions, define scope clearly, and get approval before implementation.
3. **Actionability:** Each task should be concrete, ordered logically, and include testing.

### PRD Guidelines
- Be explicit and jargon-free
- Focus on the "what" and "why," let developers figure out the "how"
- Don't start implementing—PRD is a planning artifact only

**Example:** Instead of "Add search," write "Add a real-time customer search on the main form that filters by name or ID, with results showing name, phone, and account status."

### Task List Guidelines
- Use concrete, achievable titles
- Order sub-tasks logically (dependencies flow naturally)
- Include testing within task groups, not as an afterthought
- Name exact files, modules, and forms being modified

**Example:**
```
- [ ] 2.0 Build search form UI
  - [ ] 2.1 Create unbound search textbox (txtSearchInput)
  - [ ] 2.2 Bind results listbox to query (lstResults)
  - [ ] 2.3 Add Clear Results button
```

### Access-Specific Requirements
- **Naming conventions:** Follow [vba-naming.instructions.md](../../instructions/access/vba-naming.instructions.md)
- **Reserved words:** Check [access-vba-reserved-words SKILL.md](../access-vba-reserved-words/SKILL.md) for collisions
- **Testing:** Use AccessPOSH module for PowerShell test automation (COM interop)
- **Form structure:** Separate tasks for data layer, UI design, and business logic

---

## Output Files

| Phase | Output File | Location |
|-------|------------|----------|
| PRD | `prd-[feature-name].md` | `/tasks/` |
| Tasks | `tasks-[feature-name].md` | `/tasks/` |

---

## Example Workflow

**User Request:** "I want to add a customer search feature to our database"

**Step 1: PRD Phase**
- Ask clarifying questions: "Should search be real-time or button-triggered?", "Which fields?", "Should results be paginated?"
- Generate PRD with user answers
- Save as `prd-customer-search.md`

**Step 2: Task Phase**
- Analyze PRD requirements
- Generate parent tasks (Create feature branch, Design schema, Build search form, Implement search logic, Write tests)
- Wait for user "Go" confirmation
- Generate detailed sub-tasks for each parent
- Save as `tasks-customer-search.md`

**Step 3: Implementation**
- Follow task list sequentially
- Check off tasks as completed
- Run tests to verify before marking "done"
