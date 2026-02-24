import "./taskpane.css";

// ============================================================================
// CONFIGURATION
// ============================================================================

// UPDATE THIS: Replace with your Lovable/Supabase project URL
const API_BASE_URL = "https://pgrfsrqnozhhxdovqnvc.supabase.co/functions/v1";

const CITATION_STYLES = ["APA", "MLA", "Chicago", "Harvard", "IEEE"];

// ============================================================================
// TYPES
// ============================================================================

interface AuthTokens {
  access_token: string;
  refresh_token: string;
  expires_in?: number;
}

interface User {
  id: string;
  email: string;
}

interface Project {
  id: string;
  name: string;
  citation_style: string;
}

interface Reference {
  id: string;
  title: string;
  authors?: string;
  year?: number;
  source?: string;
  formatted_apa?: string;
  formatted_mla?: string;
  formatted_chicago?: string;
  formatted_harvard?: string;
  formatted_ieee?: string;
  [key: string]: string | number | undefined;
}

interface FormattedCitation {
  inline: string;
  bibliography: string;
}

interface AuthState {
  isAuthenticated: boolean;
  user: User | null;
  accessToken: string | null;
  refreshToken: string | null;
}

// ============================================================================
// UTILITIES
// ============================================================================

/** Escape HTML to prevent XSS when inserting user-provided data into innerHTML */
function escapeHtml(str: string): string {
  const div = document.createElement("div");
  div.appendChild(document.createTextNode(str));
  return div.innerHTML;
}

// ============================================================================
// AUTH SERVICE
// ============================================================================

class AuthService {
  private authState: AuthState = {
    isAuthenticated: false,
    user: null,
    accessToken: null,
    refreshToken: null,
  };

  constructor() {
    this.loadFromStorage();
  }

  private loadFromStorage(): void {
    try {
      const stored = sessionStorage.getItem("lably_auth_state");
      if (stored) {
        this.authState = JSON.parse(stored);
      }
    } catch (e) {
      console.error("Failed to load auth state:", e);
    }
  }

  private saveToStorage(): void {
    sessionStorage.setItem("lably_auth_state", JSON.stringify(this.authState));
  }

  async login(email: string, password: string): Promise<boolean> {
    const response = await fetch(`${API_BASE_URL}/office-auth`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ action: "login", email, password }),
    });

    if (!response.ok) {
      const error = await response.json().catch(() => ({ message: "Login failed" }));
      throw new Error(error.message || "Login failed");
    }

    const data = await response.json();

    this.authState = {
      isAuthenticated: true,
      user: data.user,
      accessToken: data.access_token,
      refreshToken: data.refresh_token,
    };

    this.saveToStorage();
    return true;
  }

  async refreshAccessToken(): Promise<boolean> {
    if (!this.authState.refreshToken) {
      this.logout();
      return false;
    }

    try {
      const response = await fetch(`${API_BASE_URL}/office-auth`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          action: "refresh",
          refresh_token: this.authState.refreshToken,
        }),
      });

      if (!response.ok) {
        this.logout();
        return false;
      }

      const data = await response.json();
      this.authState.accessToken = data.access_token;
      this.authState.refreshToken = data.refresh_token;
      this.saveToStorage();
      return true;
    } catch {
      this.logout();
      return false;
    }
  }

  logout(): void {
    this.authState = {
      isAuthenticated: false,
      user: null,
      accessToken: null,
      refreshToken: null,
    };
    sessionStorage.removeItem("lably_auth_state");
  }

  getState(): AuthState {
    return this.authState;
  }

  getAuthHeader(): Record<string, string> | null {
    if (!this.authState.accessToken) return null;
    return { Authorization: `Bearer ${this.authState.accessToken}` };
  }
}

// ============================================================================
// API SERVICE
// ============================================================================

class CitationAPI {
  constructor(private auth: AuthService) {}

  private async request<T>(endpoint: string, options: RequestInit = {}): Promise<T> {
    const authHeader = this.auth.getAuthHeader();
    if (!authHeader) throw new Error("Not authenticated");

    let response = await fetch(`${API_BASE_URL}${endpoint}`, {
      ...options,
      headers: { ...options.headers, ...authHeader },
    });

    // If 401, try to refresh token and retry once
    if (response.status === 401) {
      const refreshed = await this.auth.refreshAccessToken();
      if (refreshed) {
        const newAuthHeader = this.auth.getAuthHeader();
        response = await fetch(`${API_BASE_URL}${endpoint}`, {
          ...options,
          headers: { ...options.headers, ...newAuthHeader },
        });
      } else {
        throw new Error("Authentication required");
      }
    }

    if (!response.ok) {
      const error = await response.json().catch(() => ({ message: "API request failed" }));
      throw new Error(error.message || "API request failed");
    }

    return response.json();
  }

  async getProjects(): Promise<Project[]> {
    return this.request("/office-citations?action=projects");
  }

  async getReferences(projectId: string, search?: string, style?: string): Promise<Reference[]> {
    let url = `/office-citations?action=references&projectId=${encodeURIComponent(projectId)}`;
    if (search) url += `&search=${encodeURIComponent(search)}`;
    if (style) url += `&style=${encodeURIComponent(style)}`;
    return this.request(url);
  }

  async formatCitations(
    referenceIds: string[],
    style: string
  ): Promise<Record<string, FormattedCitation>> {
    const ids = referenceIds.join(",");
    return this.request(
      `/office-citations?action=format&referenceIds=${encodeURIComponent(ids)}&style=${encodeURIComponent(style)}`
    );
  }

  async getBibliography(
    projectId: string,
    referenceIds: string[],
    style: string
  ): Promise<{ bibliography: string }> {
    const ids = referenceIds.join(",");
    return this.request(
      `/office-citations?action=bibliography&projectId=${encodeURIComponent(projectId)}&referenceIds=${encodeURIComponent(ids)}&style=${encodeURIComponent(style)}`
    );
  }
}

// ============================================================================
// WORD DOCUMENT SERVICE
// ============================================================================

class WordService {
  async insertInlineCitation(text: string, position: "current" | "end" = "current"): Promise<void> {
    return Word.run(async (context) => {
      if (position === "end") {
        const endRange = context.document.body.getRange("End");
        endRange.insertText(text, Word.InsertLocation.after);
      } else {
        const selection = context.document.getSelection();
        selection.insertText(text, Word.InsertLocation.after);
      }
      await context.sync();
    });
  }

  async insertBibliography(text: string): Promise<void> {
    return Word.run(async (context) => {
      const body = context.document.body;
      // Add a heading for the bibliography section
      const heading = body.insertParagraph("Bibliography", Word.InsertLocation.end);
      heading.styleBuiltIn = Word.BuiltInStyleName.heading1;
      // Insert the bibliography text
      body.insertParagraph(text, Word.InsertLocation.end);
      await context.sync();
    });
  }
}

// ============================================================================
// UI CONTROLLER
// ============================================================================

class UIController {
  private auth = new AuthService();
  private api = new CitationAPI(this.auth);
  private word = new WordService();
  private currentProject: Project | null = null;
  private selectedReferences: Set<string> = new Set();
  private currentStyle = "APA";
  private searchDebounceTimer: ReturnType<typeof setTimeout> | null = null;

  async initialize(): Promise<void> {
    await Office.onReady();

    const authState = this.auth.getState();
    if (authState.isAuthenticated) {
      this.showProjectView();
    } else {
      this.showLoginView();
    }
  }

  // ========================================================================
  // LOGIN VIEW
  // ========================================================================

  private showLoginView(): void {
    const root = document.getElementById("root")!;
    root.innerHTML = `
      <div class="login-container">
        <h2>Lably</h2>
        <p class="subtitle">Sign in to manage your citations</p>
        <form id="loginForm" class="login-form">
          <div class="form-group">
            <label for="email">Email</label>
            <input type="email" id="email" name="email" required placeholder="you@example.com" />
          </div>
          <div class="form-group">
            <label for="password">Password</label>
            <input type="password" id="password" name="password" required placeholder="Enter your password" />
          </div>
          <button type="submit" class="btn btn-primary" id="loginBtn">Sign In</button>
        </form>
        <div id="loginError" class="error-message" style="display: none;"></div>
      </div>
    `;

    document.getElementById("loginForm")!.addEventListener("submit", (e) => this.handleLogin(e));
  }

  private async handleLogin(e: Event): Promise<void> {
    e.preventDefault();

    const email = (document.getElementById("email") as HTMLInputElement).value;
    const password = (document.getElementById("password") as HTMLInputElement).value;
    const errorDiv = document.getElementById("loginError")!;
    const loginBtn = document.getElementById("loginBtn") as HTMLButtonElement;

    try {
      errorDiv.style.display = "none";
      loginBtn.disabled = true;
      loginBtn.textContent = "Signing in...";

      await this.auth.login(email, password);
      this.showProjectView();
    } catch (error) {
      errorDiv.textContent = error instanceof Error ? error.message : "Login failed";
      errorDiv.style.display = "block";
      loginBtn.disabled = false;
      loginBtn.textContent = "Sign In";
    }
  }

  // ========================================================================
  // PROJECT VIEW
  // ========================================================================

  private async showProjectView(): Promise<void> {
    const root = document.getElementById("root")!;
    root.innerHTML = `
      <div class="projects-container">
        <div class="header">
          <h2>Your Projects</h2>
          <button id="logoutBtn" class="btn btn-secondary btn-small">Logout</button>
        </div>
        <div id="projectsList" class="projects-list"></div>
        <div id="loadingProjects" class="loading">Loading projects...</div>
      </div>
    `;

    document.getElementById("logoutBtn")!.addEventListener("click", () => {
      this.auth.logout();
      this.showLoginView();
    });

    try {
      const projects = await this.api.getProjects();
      const projectsList = document.getElementById("projectsList")!;
      const loadingDiv = document.getElementById("loadingProjects")!;
      loadingDiv.style.display = "none";

      if (projects.length === 0) {
        projectsList.innerHTML = `
          <div class="empty-state">
            <p>No projects found. Create a project in Lably to get started.</p>
          </div>
        `;
        return;
      }

      projectsList.innerHTML = projects
        .map(
          (project) => `
          <div class="project-card" data-project-id="${escapeHtml(project.id)}">
            <h3>${escapeHtml(project.name)}</h3>
            <p class="citation-style">Style: ${escapeHtml(project.citation_style)}</p>
            <button class="btn btn-primary" data-action="select" data-project-id="${escapeHtml(project.id)}">
              Select Project
            </button>
          </div>
        `
        )
        .join("");

      document.querySelectorAll<HTMLElement>('[data-action="select"]').forEach((btn) => {
        btn.addEventListener("click", () => {
          const projectId = btn.getAttribute("data-project-id")!;
          const project = projects.find((p) => p.id === projectId)!;
          this.showReferencesView(project);
        });
      });
    } catch (error) {
      const loadingDiv = document.getElementById("loadingProjects")!;
      loadingDiv.textContent = error instanceof Error ? error.message : "Failed to load projects";
      loadingDiv.classList.add("error");
    }
  }

  // ========================================================================
  // REFERENCES VIEW
  // ========================================================================

  private async showReferencesView(project: Project): Promise<void> {
    this.currentProject = project;
    this.selectedReferences.clear();

    const root = document.getElementById("root")!;
    root.innerHTML = `
      <div class="references-container">
        <div class="header">
          <button id="backBtn" class="btn btn-secondary btn-small">Back</button>
          <h2>${escapeHtml(project.name)}</h2>
        </div>
        <div class="controls">
          <input type="text" id="searchInput" class="search-input" placeholder="Search references..." />
          <div class="style-selector">
            <label for="styleSelect">Citation Style:</label>
            <select id="styleSelect" class="style-select">
              ${CITATION_STYLES.map(
                (s) => `<option value="${s}" ${s === (project.citation_style || "APA") ? "selected" : ""}>${s}</option>`
              ).join("")}
            </select>
          </div>
        </div>
        <div id="referencesList" class="references-list"></div>
        <div id="loadingReferences" class="loading">Loading references...</div>
        <div class="action-bar">
          <button id="insertInlineBtn" class="btn btn-primary" disabled>Insert Citation</button>
          <button id="insertBibBtn" class="btn btn-secondary" disabled>Insert Bibliography</button>
        </div>
      </div>
    `;

    this.currentStyle = project.citation_style || "APA";

    document.getElementById("backBtn")!.addEventListener("click", () => this.showProjectView());

    document.getElementById("styleSelect")!.addEventListener("change", (e) => {
      this.currentStyle = (e.target as HTMLSelectElement).value;
      this.loadReferences();
    });

    document.getElementById("searchInput")!.addEventListener("input", () => {
      // Debounce search: wait 300ms after user stops typing
      if (this.searchDebounceTimer) clearTimeout(this.searchDebounceTimer);
      this.searchDebounceTimer = setTimeout(() => this.loadReferences(), 300);
    });

    document.getElementById("insertInlineBtn")!.addEventListener("click", () => {
      this.insertSelectedCitations();
    });

    document.getElementById("insertBibBtn")!.addEventListener("click", () => {
      this.insertBibliography();
    });

    this.loadReferences();
  }

  private async loadReferences(): Promise<void> {
    if (!this.currentProject) return;

    const referencesList = document.getElementById("referencesList");
    const loadingDiv = document.getElementById("loadingReferences");
    if (!referencesList || !loadingDiv) return;

    loadingDiv.style.display = "block";
    loadingDiv.textContent = "Loading references...";
    loadingDiv.classList.remove("error");

    try {
      const searchTerm = (document.getElementById("searchInput") as HTMLInputElement)?.value || "";
      const references = await this.api.getReferences(
        this.currentProject.id,
        searchTerm,
        this.currentStyle
      );

      loadingDiv.style.display = "none";

      if (references.length === 0) {
        referencesList.innerHTML = `
          <div class="empty-state">
            <p>No references found.</p>
          </div>
        `;
        return;
      }

      const styleKey = `formatted_${this.currentStyle.toLowerCase()}`;

      referencesList.innerHTML = references
        .map(
          (ref) => `
          <div class="reference-item" data-ref-id="${escapeHtml(ref.id)}">
            <input
              type="checkbox"
              class="ref-checkbox"
              data-ref-id="${escapeHtml(ref.id)}"
              ${this.selectedReferences.has(ref.id) ? "checked" : ""}
            />
            <div class="ref-content">
              <h4>${escapeHtml(ref.title)}</h4>
              <p class="ref-authors">${escapeHtml(ref.authors || "No authors")}</p>
              <p class="ref-year">${ref.year || ""}</p>
              <p class="ref-formatted">${escapeHtml(String(ref[styleKey] || ""))}</p>
            </div>
          </div>
        `
        )
        .join("");

      document.querySelectorAll<HTMLInputElement>(".ref-checkbox").forEach((checkbox) => {
        checkbox.addEventListener("change", () => {
          const refId = checkbox.getAttribute("data-ref-id")!;
          if (checkbox.checked) {
            this.selectedReferences.add(refId);
          } else {
            this.selectedReferences.delete(refId);
          }
          this.updateActionButtons();
        });
      });

      this.updateActionButtons();
    } catch (error) {
      loadingDiv.textContent = error instanceof Error ? error.message : "Failed to load references";
      loadingDiv.classList.add("error");
      loadingDiv.style.display = "block";
    }
  }

  private updateActionButtons(): void {
    const insertInlineBtn = document.getElementById("insertInlineBtn") as HTMLButtonElement | null;
    const insertBibBtn = document.getElementById("insertBibBtn") as HTMLButtonElement | null;
    const hasSelected = this.selectedReferences.size > 0;

    if (insertInlineBtn) insertInlineBtn.disabled = !hasSelected;
    if (insertBibBtn) insertBibBtn.disabled = !hasSelected;
  }

  private async insertSelectedCitations(): Promise<void> {
    if (this.selectedReferences.size === 0) return;

    try {
      const refIds = Array.from(this.selectedReferences);
      const formatted = await this.api.formatCitations(refIds, this.currentStyle);

      for (const refId of refIds) {
        const citation = formatted[refId];
        if (citation) {
          await this.word.insertInlineCitation(citation.inline);
        }
      }
    } catch (error) {
      alert(error instanceof Error ? error.message : "Failed to insert citations");
    }
  }

  private async insertBibliography(): Promise<void> {
    if (!this.currentProject || this.selectedReferences.size === 0) return;

    try {
      const refIds = Array.from(this.selectedReferences);
      const result = await this.api.getBibliography(
        this.currentProject.id,
        refIds,
        this.currentStyle
      );

      await this.word.insertBibliography(result.bibliography);
    } catch (error) {
      alert(error instanceof Error ? error.message : "Failed to insert bibliography");
    }
  }
}

// ============================================================================
// INITIALIZATION
// ============================================================================

const uiController = new UIController();
uiController.initialize().catch((error) => {
  console.error("Failed to initialize Lably add-in:", error);
  const root = document.getElementById("root");
  if (root) {
    root.innerHTML = `
      <div class="error-container">
        <h2>Error</h2>
        <p>Failed to initialize the Lably add-in.</p>
        <p>${error instanceof Error ? escapeHtml(error.message) : "Unknown error"}</p>
      </div>
    `;
  }
});
