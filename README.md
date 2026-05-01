# RelatedEntityListAsTagControl

A Dynamics 365 **PowerApps Component Framework (PCF)** dataset control that renders a one-to-many related entity subgrid as a **tagging interface**.  
Users can search for related records via a keyword search / drop-down menu and add or remove associations using the Dataverse associate / disassociate WebAPI.

---

## Features

- Displays currently associated records as removable tag pills
- Keyword search with a drop-down result list (300 ms debounce)
- Adds associations via `WebAPI.associateRecord`
- Removes associations via `WebAPI.disassociateRecord`
- Refreshes the underlying dataset after each change
- Fluent UI–inspired styling that matches the Dynamics 365 look-and-feel
- Accessible markup (`role`, `aria-label`, `aria-live`)
- Error banner for user-friendly failure messages

---

## Prerequisites

| Tool | Version |
|------|---------|
| Node.js | 16 + |
| npm | 8 + |
| [Power Platform CLI (`pac`)](https://learn.microsoft.com/en-us/power-platform/developer/cli/introduction) | latest |

---

## Project structure

```
RelatedEntityListAsTagControl/
├── ControlManifest.Input.xml   # PCF manifest – declares properties & resources
├── index.ts                    # TypeScript control implementation
├── css/
│   └── RelatedEntityTagControl.css
├── generated/
│   └── ManifestTypes.d.ts      # Auto-generated type definitions
├── package.json
└── tsconfig.json
```

---

## Build

```bash
# Install dependencies
npm install

# Development build (watch mode with test harness)
npm start

# Production build
npm run build
```

The compiled output is written to the `out/` folder.

---

## Configuration properties

| Property | Type | Required | Description |
|---|---|---|---|
| `PrimaryFieldName` | `SingleLine.Text` | ✅ | Logical name of the field on the related entity to display as the tag label (e.g. `name`, `fullname`). |
| `SearchFields` | `SingleLine.Text` | ✅ | Comma-separated list of field names to search against (e.g. `name,emailaddress1`). |
| `RelatedEntityName` | `SingleLine.Text` | ✅ | Logical name of the related entity (e.g. `contact`, `account`). |
| `RelationshipName` | `SingleLine.Text` | ✅ | Schema name of the relationship used for associate/disassociate (e.g. `account_contacts`). |

---

## Deployment

1. Build the solution:
   ```bash
   npm run build
   ```

2. Pack a solution using the Power Platform CLI:
   ```bash
   pac solution init --publisher-name xrmatic --publisher-prefix xrm
   pac solution add-reference --path .
   pac solution pack --folder . --packagetype Managed
   ```

3. Import the generated `.zip` file into your Dynamics 365 / Power Platform environment.

4. Open a model-driven app form, add a subgrid bound to the desired relationship, and set **Custom Control** to `RelatedEntityTagControl`.

5. Configure the four control properties (see table above) to match the relationship and entity you want to tag.

---

## Usage

Once deployed and configured on a form:

- **Existing associations** appear automatically as tag pills inside the control.
- **Search** – start typing in the search box; a dropdown lists matching records (excluding records already tagged).
- **Add** a record – click its entry in the dropdown.
- **Remove** a record – click the `×` on its tag pill.

All changes call the Dataverse WebAPI and refresh the dataset immediately.
