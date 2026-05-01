import { IInputs, IOutputs } from "./generated/ManifestTypes";

/**
 * RelatedEntityTagControl
 *
 * A Dynamics 365 PCF dataset control that renders a one-to-many related entity
 * subgrid as a tagging interface.  Users can search for related entity records
 * via a keyword search / drop-down menu and add or remove associations using
 * the Dataverse associate / disassociate WebAPI.
 */
export class RelatedEntityTagControl
    implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    // ── PCF context ──────────────────────────────────────────────────────────
    private _context!: ComponentFramework.Context<IInputs>;
    private _container!: HTMLDivElement;

    // ── DOM elements ─────────────────────────────────────────────────────────
    private _tagContainer!: HTMLDivElement;
    private _searchInput!: HTMLInputElement;
    private _dropdown!: HTMLDivElement;
    private _loadingIndicator!: HTMLDivElement;
    private _errorBanner!: HTMLDivElement;

    // ── State ─────────────────────────────────────────────────────────────────
    /** id → display name for currently associated records */
    private _associatedRecords: Map<string, string> = new Map();
    /** Results returned by the last keyword search */
    private _searchResults: Array<{ id: string; name: string }> = [];
    /** Debounce timer for the search input */
    private _searchDebounce: ReturnType<typeof setTimeout> | null = null;

    // ─────────────────────────────────────────────────────────────────────────
    // Lifecycle
    // ─────────────────────────────────────────────────────────────────────────

    public init(
        context: ComponentFramework.Context<IInputs>,
        _notifyOutputChanged: () => void,
        _state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this._context = context;
        this._container = container;
        this._buildUI();
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._context = context;
        const dataset = context.parameters.relatedRecords;

        if (!dataset || dataset.loading) return;

        const primaryField =
            (context.parameters.PrimaryFieldName.raw ?? "").trim() || "name";

        this._associatedRecords.clear();
        for (const id of dataset.sortedRecordIds) {
            const record = dataset.records[id];
            const displayName =
                record.getFormattedValue(primaryField) ||
                record.getFormattedValue("name") ||
                id;
            this._associatedRecords.set(id, displayName);
        }

        this._renderTags();
    }

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void {
        if (this._searchDebounce !== null) {
            clearTimeout(this._searchDebounce);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // UI construction
    // ─────────────────────────────────────────────────────────────────────────

    private _buildUI(): void {
        this._container.className = "tag-control-container";

        // ── Input wrapper ────────────────────────────────────────────────────
        const inputWrapper = document.createElement("div");
        inputWrapper.className = "tag-input-wrapper";

        this._tagContainer = document.createElement("div");
        this._tagContainer.className = "tag-container";

        this._searchInput = document.createElement("input");
        this._searchInput.type = "text";
        this._searchInput.className = "tag-search-input";
        this._searchInput.placeholder = "Search to add…";
        this._searchInput.setAttribute("aria-label", "Search for records to add");
        this._searchInput.addEventListener("input", this._onSearchInput.bind(this));
        this._searchInput.addEventListener("keydown", this._onKeyDown.bind(this));
        this._searchInput.addEventListener("blur", this._onBlur.bind(this));

        inputWrapper.appendChild(this._tagContainer);
        inputWrapper.appendChild(this._searchInput);

        // ── Dropdown ──────────────────────────────────────────────────────────
        this._dropdown = document.createElement("div");
        this._dropdown.className = "tag-dropdown hidden";
        this._dropdown.setAttribute("role", "listbox");

        // ── Loading indicator ─────────────────────────────────────────────────
        this._loadingIndicator = document.createElement("div");
        this._loadingIndicator.className = "tag-loading hidden";
        this._loadingIndicator.setAttribute("aria-live", "polite");
        this._loadingIndicator.textContent = "Searching…";

        // ── Error banner ──────────────────────────────────────────────────────
        this._errorBanner = document.createElement("div");
        this._errorBanner.className = "tag-error hidden";
        this._errorBanner.setAttribute("role", "alert");

        this._container.appendChild(inputWrapper);
        this._container.appendChild(this._loadingIndicator);
        this._container.appendChild(this._dropdown);
        this._container.appendChild(this._errorBanner);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Tag rendering
    // ─────────────────────────────────────────────────────────────────────────

    private _renderTags(): void {
        this._tagContainer.innerHTML = "";
        this._associatedRecords.forEach((name, id) => {
            this._tagContainer.appendChild(this._createTag(id, name));
        });
    }

    private _createTag(id: string, name: string): HTMLSpanElement {
        const tag = document.createElement("span");
        tag.className = "tag";
        tag.setAttribute("role", "listitem");
        tag.dataset.id = id;

        const label = document.createElement("span");
        label.className = "tag-label";
        label.textContent = name;
        label.title = name;

        const removeBtn = document.createElement("button");
        removeBtn.className = "tag-remove";
        removeBtn.textContent = "×";
        removeBtn.title = `Remove ${name}`;
        removeBtn.setAttribute("aria-label", `Remove ${name}`);
        removeBtn.addEventListener("click", () => this._removeTag(id));

        tag.appendChild(label);
        tag.appendChild(removeBtn);
        return tag;
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Search input handlers
    // ─────────────────────────────────────────────────────────────────────────

    private _onSearchInput(): void {
        const value = this._searchInput.value.trim();

        if (this._searchDebounce !== null) {
            clearTimeout(this._searchDebounce);
            this._searchDebounce = null;
        }

        if (value.length === 0) {
            this._hideDropdown();
            return;
        }

        this._searchDebounce = setTimeout(() => {
            void this._performSearch(value);
        }, 300);
    }

    private _onKeyDown(event: KeyboardEvent): void {
        if (event.key === "Escape") {
            this._hideDropdown();
            this._searchInput.value = "";
        }
    }

    private _onBlur(): void {
        // Delay so that mousedown on a dropdown item fires before the dropdown hides.
        setTimeout(() => this._hideDropdown(), 200);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Search / dropdown
    // ─────────────────────────────────────────────────────────────────────────

    private async _performSearch(query: string): Promise<void> {
        const entityName =
            (this._context.parameters.RelatedEntityName.raw ?? "").trim();
        const searchFieldsRaw =
            (this._context.parameters.SearchFields.raw ?? "").trim() || "name";
        const primaryField =
            (this._context.parameters.PrimaryFieldName.raw ?? "").trim() || "name";

        if (!entityName) {
            this._showError(
                "RelatedEntityName is not configured. Please set the control property."
            );
            return;
        }

        this._clearError();
        this._showLoading();

        try {
            const searchFields = searchFieldsRaw
                .split(",")
                .map((f) => f.trim())
                .filter((f) => f.length > 0);

            // Build OData $filter using `contains` on each search field
            const safeQuery = query.replace(/'/g, "''");
            const filterParts = searchFields.map(
                (f) => `contains(${f},'${safeQuery}')`
            );
            const filter = filterParts.join(" or ");

            // $select: union of search fields + primary field
            const selectFields = [
                ...new Set([...searchFields, primaryField]),
            ].join(",");

            const result =
                await this._context.webAPI.retrieveMultipleRecords(
                    entityName,
                    `?$select=${selectFields}&$filter=${filter}&$top=10`
                );

            // Dataverse primary key convention: {entityname}id
            const idField = `${entityName}id`;

            this._searchResults = result.entities.map((entity) => ({
                id: (entity[idField] as string) ?? (entity["id"] as string) ?? "",
                name:
                    (entity[primaryField] as string) ??
                    (entity["name"] as string) ??
                    "",
            }));

            this._renderDropdown();
        } catch (err) {
            this._hideLoading();
            const message =
                err instanceof Error ? err.message : "An unexpected error occurred.";
            this._showError(`Search failed: ${message}`);
        }
    }

    private _renderDropdown(): void {
        this._hideLoading();
        this._dropdown.innerHTML = "";

        const unassociated = this._searchResults.filter(
            (r) => r.id && !this._associatedRecords.has(r.id)
        );

        if (unassociated.length === 0) {
            const noResults = document.createElement("div");
            noResults.className = "tag-dropdown-item no-results";
            noResults.textContent =
                this._searchResults.length === 0
                    ? "No results found"
                    : "All matching records are already added";
            this._dropdown.appendChild(noResults);
        } else {
            for (const result of unassociated) {
                const item = document.createElement("div");
                item.className = "tag-dropdown-item";
                item.textContent = result.name;
                item.setAttribute("role", "option");
                item.addEventListener("mousedown", (e) => {
                    // Prevent the input's blur from firing before the click
                    e.preventDefault();
                    void this._addTag(result.id, result.name);
                });
                this._dropdown.appendChild(item);
            }
        }

        this._dropdown.classList.remove("hidden");
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Associate / disassociate
    // ─────────────────────────────────────────────────────────────────────────

    /**
     * Extended WebApi interface that includes associate/disassociate methods.
     * These exist at runtime in Dynamics 365 but are not declared in the
     * @types/powerapps-component-framework 1.3.x typings.
     */
    private get _webApi(): ComponentFramework.WebApi & {
        associateRecord(
            entityType: string,
            entityId: string,
            relationship: string,
            relatedEntityId: string,
            relatedEntityType: string
        ): Promise<void>;
        disassociateRecord(
            entityType: string,
            entityId: string,
            relationship: string,
            relatedEntityId: string
        ): Promise<void>;
    } {
        return this._context.webAPI as ComponentFramework.WebApi & {
            associateRecord(
                entityType: string,
                entityId: string,
                relationship: string,
                relatedEntityId: string,
                relatedEntityType: string
            ): Promise<void>;
            disassociateRecord(
                entityType: string,
                entityId: string,
                relationship: string,
                relatedEntityId: string
            ): Promise<void>;
        };
    }

    private async _addTag(relatedEntityId: string, name: string): Promise<void> {
        const { entityName, relationshipName, parentEntityId, parentEntityName } =
            this._getRelationshipConfig();

        if (!entityName || !relationshipName || !parentEntityId || !parentEntityName) {
            this._showError(
                "Control is not fully configured or is not placed on a record form."
            );
            return;
        }

        try {
            await this._webApi.associateRecord(
                parentEntityName,
                parentEntityId,
                relationshipName,
                relatedEntityId,
                entityName
            );

            this._associatedRecords.set(relatedEntityId, name);
            this._searchInput.value = "";
            this._hideDropdown();
            this._renderTags();
            this._context.parameters.relatedRecords.refresh();
            this._clearError();
        } catch (err) {
            const message =
                err instanceof Error ? err.message : "An unexpected error occurred.";
            this._showError(`Failed to add tag: ${message}`);
        }
    }

    private async _removeTag(relatedEntityId: string): Promise<void> {
        const { entityName, relationshipName, parentEntityId, parentEntityName } =
            this._getRelationshipConfig();

        if (!entityName || !relationshipName || !parentEntityId || !parentEntityName) {
            this._showError(
                "Control is not fully configured or is not placed on a record form."
            );
            return;
        }

        try {
            await this._webApi.disassociateRecord(
                parentEntityName,
                parentEntityId,
                relationshipName,
                relatedEntityId
            );

            this._associatedRecords.delete(relatedEntityId);
            this._renderTags();
            this._context.parameters.relatedRecords.refresh();
            this._clearError();
        } catch (err) {
            const message =
                err instanceof Error ? err.message : "An unexpected error occurred.";
            this._showError(`Failed to remove tag: ${message}`);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Helpers
    // ─────────────────────────────────────────────────────────────────────────

    private _getRelationshipConfig(): {
        entityName: string;
        relationshipName: string;
        parentEntityId: string;
        parentEntityName: string;
    } {
        const entityName =
            (this._context.parameters.RelatedEntityName.raw ?? "").trim();
        const relationshipName =
            (this._context.parameters.RelationshipName.raw ?? "").trim();

        // `context.page` exposes the host form's record information
        const page = (this._context as unknown as {
            page: { entityId: string; entityTypeName: string };
        }).page;

        return {
            entityName,
            relationshipName,
            parentEntityId: page?.entityId ?? "",
            parentEntityName: page?.entityTypeName ?? "",
        };
    }

    private _showLoading(): void {
        this._loadingIndicator.classList.remove("hidden");
        this._dropdown.classList.add("hidden");
    }

    private _hideLoading(): void {
        this._loadingIndicator.classList.add("hidden");
    }

    private _hideDropdown(): void {
        this._dropdown.classList.add("hidden");
    }

    private _showError(message: string): void {
        this._errorBanner.textContent = message;
        this._errorBanner.classList.remove("hidden");
    }

    private _clearError(): void {
        this._errorBanner.textContent = "";
        this._errorBanner.classList.add("hidden");
    }
}
