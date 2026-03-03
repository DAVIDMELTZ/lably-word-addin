import "./taskpane.css";

// ============================================================================
// CONFIGURATION
// ============================================================================

const LABLY_VERSION = "v7";
const API_BASE_URL = "https://pgrfsrqnozhhxdovqnvc.supabase.co/functions/v1";
// Local fallback styles — used only when backend doesn't provide a styles list
const FALLBACK_STYLES = ["APA", "MLA", "Chicago", "Harvard", "IEEE"];

// ============================================================================
// TYPES
// ============================================================================

interface User { id: string; email: string; }

interface Project {
  id: string;
  name: string;
  citation_style: string;
  [key: string]: any;
}

interface Reference {
  id: string;
  title: string;
  authors?: any; // Can be string, array of strings, or array of objects
  year?: number;
  source?: string;
  [key: string]: any;
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

function escapeHtml(str: string): string {
  const div = document.createElement("div");
  div.appendChild(document.createTextNode(str));
  return div.innerHTML;
}

function extractArray(response: any): any[] {
  if (Array.isArray(response)) return response;
  if (response && typeof response === "object") {
    for (const key of ["data", "projects", "references", "items", "results", "citations"]) {
      if (Array.isArray(response[key])) return response[key];
    }
    for (const key of Object.keys(response)) {
      if (Array.isArray(response[key])) return response[key];
    }
  }
  return [];
}

// ============================================================================
// AUTHOR NAME EXTRACTION & CITATION FORMATTING
// ============================================================================

interface AuthorName {
  first: string;
  last: string;
}

/** Parse a single name string into first/last parts. */
function parseName(name: string): AuthorName {
  const trimmed = name.trim();
  if (!trimmed) return { first: "", last: "Unknown" };
  // "Last, First" format
  if (trimmed.includes(",")) {
    const parts = trimmed.split(",");
    return { last: parts[0].trim(), first: parts.slice(1).join(",").trim() };
  }
  // "First Last" format
  const parts = trimmed.split(/\s+/);
  if (parts.length === 1) return { first: "", last: parts[0] };
  return { first: parts.slice(0, -1).join(" "), last: parts[parts.length - 1] };
}

/** Extract author name from an object with various possible field names. */
function parseAuthorObject(obj: any): AuthorName {
  if (!obj || typeof obj !== "object") return { first: "", last: "Unknown" };
  const first = obj.first_name || obj.firstName || obj.given || "";
  const last = obj.last_name || obj.lastName || obj.family || obj.surname || "";
  if (first || last) return { first: String(first), last: String(last) || "Unknown" };
  const fullName = obj.full_name || obj.fullName || obj.name || "";
  if (fullName) return parseName(String(fullName));
  return { first: "", last: "Unknown" };
}

/** Extract a structured array of author names from any format the API returns. */
function extractAuthors(authors: any): AuthorName[] {
  if (!authors) return [];
  if (typeof authors === "string") {
    return authors.split(/\s*[;&]\s*/).map(parseName).filter(a => a.last !== "Unknown" || a.first);
  }
  if (Array.isArray(authors)) {
    return authors.map((a: any) => {
      if (typeof a === "string") return parseName(a);
      if (typeof a === "object") return parseAuthorObject(a);
      return { first: "", last: String(a) };
    }).filter(a => a.last);
  }
  if (typeof authors === "object") return [parseAuthorObject(authors)];
  return [{ first: "", last: String(authors) }];
}

/** Get initials from a first name: "Christina Nicole" → "C. N." */
function getInitials(firstName: string): string {
  if (!firstName) return "";
  return firstName.split(/[\s.]+/).filter(Boolean).map(n => n.charAt(0).toUpperCase() + ".").join(" ");
}

/** Convert authors to a human-readable display string (for the UI list). */
function authorsToString(authors: any): string {
  const parsed = extractAuthors(authors);
  if (parsed.length === 0) return "Unknown";
  return parsed.map(a => {
    if (a.first && a.last) return `${a.first} ${a.last}`;
    return a.last || a.first || "Unknown";
  }).join(", ");
}

/**
 * Build an inline citation from reference data, formatted per the selected style.
 *
 * APA 7th:    1 author: (Smith, 2023)   2: (Smith & Doe, 2023)    3+: (Smith et al., 2023)
 * MLA 9th:    1: (Smith)                 2: (Smith and Doe)        3+: (Smith et al.)          — no year
 * Chicago:    1: (Smith 2023)            2: (Smith and Doe 2023)   3+: (Smith et al. 2023)     — no comma before year
 * Harvard:    1: (Smith, 2023)           2: (Smith and Doe, 2023)  3: (Smith, Doe and Jones, 2023)  4+: (Smith et al., 2023)
 * IEEE:       numbered: [1], [2], etc.   (ieeeNumber must be set by caller)
 */
function buildInlineCitation(ref: Reference, style: string, ieeeNumber?: number): string {
  const authors = extractAuthors(ref.authors);
  const year = ref.year || "n.d.";
  const n = authors.length;
  const last1 = n > 0 ? authors[0].last : "Unknown";

  switch (style.toUpperCase()) {
    case "APA":
      if (n <= 1) return `(${last1}, ${year})`;
      if (n === 2) return `(${last1} & ${authors[1].last}, ${year})`;
      return `(${last1} et al., ${year})`;

    case "MLA":
      if (n <= 1) return `(${last1})`;
      if (n === 2) return `(${last1} and ${authors[1].last})`;
      return `(${last1} et al.)`;

    case "CHICAGO":
      if (n <= 1) return `(${last1} ${year})`;
      if (n === 2) return `(${last1} and ${authors[1].last} ${year})`;
      return `(${last1} et al. ${year})`;

    case "HARVARD":
      if (n <= 1) return `(${last1}, ${year})`;
      if (n === 2) return `(${last1} and ${authors[1].last}, ${year})`;
      if (n === 3) return `(${last1}, ${authors[1].last} and ${authors[2].last}, ${year})`;
      return `(${last1} et al., ${year})`;

    case "IEEE":
      return `[${ieeeNumber ?? "?"}]`;

    default:
      return `(${last1}, ${year})`;
  }
}

/**
 * Build a bibliography entry formatted per the selected style.
 * Uses pre-formatted fields from the API if available, otherwise builds locally.
 *
 * APA:     Last, F. M., Last, F. M., & Last, F. M. (Year). Title. Source.
 * MLA:     Last, First. "Title." Source, Year.            (3+ authors: first et al.)
 * Chicago: Last, First, and First Last. Year. "Title." Source.
 * Harvard: Last, F. and Last, F. (Year) 'Title', Source.
 * IEEE:    [#] F. Last, "Title," Source, Year.
 */
function getBibEntry(ref: Reference, style: string, ieeeNumber?: number): string {
  // Prefer pre-formatted entry from the backend
  const styleKey = `formatted_${style.toLowerCase()}`;
  if (ref[styleKey]) return String(ref[styleKey]);

  const authors = extractAuthors(ref.authors);
  const year = ref.year || "n.d.";
  const title = ref.title || "Untitled";
  const source = ref.source || "";

  switch (style.toUpperCase()) {
    case "APA": {
      // Last, F. M., Last, F. M., & Last, F. M. (Year). Title. Source.
      const parts = authors.map(a => `${a.last}${a.first ? ", " + getInitials(a.first) : ""}`);
      let authorStr: string;
      if (parts.length === 0) authorStr = "Unknown";
      else if (parts.length === 1) authorStr = parts[0];
      else if (parts.length === 2) authorStr = `${parts[0]}, & ${parts[1]}`;
      else authorStr = parts.slice(0, -1).join(", ") + ", & " + parts[parts.length - 1];
      return `${authorStr} (${year}). ${title}.${source ? " " + source + "." : ""}`;
    }

    case "MLA": {
      // Last, First. "Title." Source, Year.  (3+ → first author et al.)
      let authorStr: string;
      if (authors.length === 0) authorStr = "Unknown";
      else if (authors.length === 1) authorStr = `${authors[0].last}, ${authors[0].first}`;
      else if (authors.length === 2) authorStr = `${authors[0].last}, ${authors[0].first}, and ${authors[1].first} ${authors[1].last}`;
      else authorStr = `${authors[0].last}, ${authors[0].first}, et al.`;
      return `${authorStr}. "${title}."${source ? " " + source + "," : ""} ${year}.`;
    }

    case "CHICAGO": {
      // Last, First, and First Last. Year. "Title." Source.
      let authorStr: string;
      if (authors.length === 0) authorStr = "Unknown";
      else if (authors.length === 1) authorStr = `${authors[0].last}, ${authors[0].first}`;
      else {
        const rest = authors.slice(1).map(a => `${a.first} ${a.last}`);
        authorStr = `${authors[0].last}, ${authors[0].first}, ` +
          (rest.length === 1 ? `and ${rest[0]}` : rest.slice(0, -1).join(", ") + ", and " + rest[rest.length - 1]);
      }
      return `${authorStr}. ${year}. "${title}."${source ? " " + source + "." : ""}`;
    }

    case "HARVARD": {
      // Last, F. and Last, F. (Year) 'Title', Source.
      const parts = authors.map(a => `${a.last}, ${a.first ? getInitials(a.first) : ""}`);
      let authorStr: string;
      if (parts.length === 0) authorStr = "Unknown";
      else if (parts.length === 1) authorStr = parts[0];
      else if (parts.length === 2) authorStr = `${parts[0]} and ${parts[1]}`;
      else authorStr = parts.slice(0, -1).join(", ") + " and " + parts[parts.length - 1];
      return `${authorStr} (${year}) '${title}',${source ? " " + source + "." : ""}`;
    }

    case "IEEE": {
      // [#] F. Last, "Title," Source, Year.
      const parts = authors.map(a => `${a.first ? getInitials(a.first) + " " : ""}${a.last}`);
      let authorStr: string;
      if (parts.length === 0) authorStr = "Unknown";
      else if (parts.length <= 6) {
        authorStr = parts.length === 1 ? parts[0] :
          parts.slice(0, -1).join(", ") + ", and " + parts[parts.length - 1];
      } else {
        authorStr = parts[0] + " et al.";
      }
      const num = ieeeNumber ?? "?";
      return `[${num}] ${authorStr}, "${title},"${source ? " " + source + "," : ""} ${year}.`;
    }

    default: {
      const authorStr = authors.map(a => `${a.last}${a.first ? ", " + getInitials(a.first) : ""}`).join(", ");
      return `${authorStr || "Unknown"} (${year}). ${title}.${source ? " " + source + "." : ""}`;
    }
  }
}

// ============================================================================
// AUTH SERVICE
// ============================================================================

class AuthService {
  private authState: AuthState = {
    isAuthenticated: false, user: null, accessToken: null, refreshToken: null,
  };

  constructor() { this.loadFromStorage(); }

  private loadFromStorage(): void {
    try {
      const stored = sessionStorage.getItem("lably_auth_state");
      if (stored) this.authState = JSON.parse(stored);
    } catch (e) { console.error("Failed to load auth state:", e); }
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
      isAuthenticated: true, user: data.user,
      accessToken: data.access_token, refreshToken: data.refresh_token,
    };
    this.saveToStorage();
    return true;
  }

  async refreshAccessToken(): Promise<boolean> {
    if (!this.authState.refreshToken) { this.logout(); return false; }
    try {
      const response = await fetch(`${API_BASE_URL}/office-auth`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "refresh", refresh_token: this.authState.refreshToken }),
      });
      if (!response.ok) { this.logout(); return false; }
      const data = await response.json();
      this.authState.accessToken = data.access_token;
      this.authState.refreshToken = data.refresh_token;
      this.saveToStorage();
      return true;
    } catch { this.logout(); return false; }
  }

  logout(): void {
    this.authState = { isAuthenticated: false, user: null, accessToken: null, refreshToken: null };
    sessionStorage.removeItem("lably_auth_state");
  }

  getState(): AuthState { return this.authState; }

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
      ...options, headers: { ...options.headers, ...authHeader },
    });

    if (response.status === 401) {
      const refreshed = await this.auth.refreshAccessToken();
      if (refreshed) {
        const newAuthHeader = this.auth.getAuthHeader();
        response = await fetch(`${API_BASE_URL}${endpoint}`, {
          ...options, headers: { ...options.headers, ...newAuthHeader },
        });
      } else { throw new Error("Session expired. Please log in again."); }
    }

    if (!response.ok) {
      const error = await response.json().catch(() => ({ message: "API request failed" }));
      throw new Error(error.message || "API request failed");
    }
    return response.json();
  }

  async getProjects(): Promise<any> {
    return this.request("/office-citations?action=projects");
  }

  /**
   * Fetch available citation styles from the backend.
   * Tries action=styles first; if it fails, returns null so the caller can
   * fall back to extracting styles from the project list.
   */
  async getStyles(): Promise<string[] | null> {
    try {
      const raw = await this.request<any>("/office-citations?action=styles");
      // Accept an array directly, or an object with a styles/data key
      if (Array.isArray(raw)) return raw.map(String);
      if (raw && typeof raw === "object") {
        for (const key of ["styles", "data", "items", "results"]) {
          if (Array.isArray(raw[key])) return raw[key].map(String);
        }
      }
      return null;
    } catch {
      // Endpoint not available — return null so caller uses fallback
      return null;
    }
  }

  async getReferences(projectId: string, search?: string, style?: string): Promise<any> {
    let url = `/office-citations?action=references&projectId=${encodeURIComponent(projectId)}`;
    if (search) url += `&search=${encodeURIComponent(search)}`;
    if (style) url += `&style=${encodeURIComponent(style)}`;
    return this.request(url);
  }

  async formatCitations(referenceIds: string[], style: string): Promise<any> {
    const ids = referenceIds.join(",");
    return this.request(
      `/office-citations?action=format&referenceIds=${encodeURIComponent(ids)}&style=${encodeURIComponent(style)}`
    );
  }

  async getBibliography(projectId: string, referenceIds: string[], style: string): Promise<any> {
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
  /**
   * Insert inline citation text at the current cursor position.
   * Uses multiple strategies with debug logging.
   */
  async insertInlineCitation(text: string): Promise<void> {
    if (!text || !text.trim()) {
      throw new Error("Citation text is empty");
    }
    const citationText = text.startsWith(" ") ? text : " " + text;

    // Strategy 1: Insert at selection using "After"
    try {
      await Word.run(async (context) => {
        const range = context.document.getSelection();
        range.insertText(citationText, "After");
        await context.sync();
      });
      return;
    } catch (e1) {
      console.warn("[Lably] Strategy 1 (selection After) failed:", e1);
    }

    // Strategy 2: Insert at selection using "Replace"
    try {
      await Word.run(async (context) => {
        const range = context.document.getSelection();
        range.insertText(citationText, "Replace");
        await context.sync();
      });
      return;
    } catch (e2) {
      console.warn("[Lably] Strategy 2 (selection Replace) failed:", e2);
    }

    // Strategy 3: Insert HTML at selection
    try {
      await Word.run(async (context) => {
        const range = context.document.getSelection();
        const safeText = citationText
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;");
        range.insertHtml("<span>" + safeText + "</span>", "After");
        await context.sync();
      });
      return;
    } catch (e3) {
      console.warn("[Lably] Strategy 3 (selection insertHtml) failed:", e3);
    }

    // Strategy 4: Append to last content paragraph directly
    try {
      await Word.run(async (context) => {
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load("items");
        await context.sync();

        // Load all paragraph texts in one batch
        for (const p of paragraphs.items) p.load("text");
        await context.sync();

        // Find last non-bibliography paragraph
        let target = paragraphs.items[0];
        for (let i = paragraphs.items.length - 1; i >= 0; i--) {
          const t = (paragraphs.items[i].text || "").trim().toLowerCase();
          if (t && t !== "bibliography" && t !== "references") {
            target = paragraphs.items[i];
            break;
          }
        }

        target.insertText(citationText, "End");
        await context.sync();
      });
      return;
    } catch (e4) {
      console.warn("[Lably] Strategy 4 (paragraph End) failed:", e4);
    }

    throw new Error("All citation insert strategies failed. Click inside the document and try again.");
  }

  /**
   * Find an existing "Bibliography" or "References" section and replace it,
   * or create one at the end of the document. Uses Normal style to match doc font.
   */
  async insertOrUpdateBibliography(entries: string[]): Promise<void> {
    return Word.run(async (context) => {
      const body = context.document.body;
      const paragraphs = body.paragraphs;
      paragraphs.load("items,text,style");
      await context.sync();

      // Look for existing "Bibliography" or "References" heading
      let bibHeadingIndex = -1;
      for (let i = 0; i < paragraphs.items.length; i++) {
        const text = paragraphs.items[i].text.trim().toLowerCase();
        if (text === "bibliography" || text === "references") {
          bibHeadingIndex = i;
          break;
        }
      }

      if (bibHeadingIndex >= 0) {
        // Delete everything from the bibliography heading to the end
        for (let i = paragraphs.items.length - 1; i >= bibHeadingIndex; i--) {
          paragraphs.items[i].delete();
        }
        await context.sync();
      }

      // Insert bibliography heading (bold, normal size — not Heading1)
      const heading = body.insertParagraph("Bibliography", Word.InsertLocation.end);
      heading.styleBuiltIn = Word.BuiltInStyleName.normal;
      heading.font.bold = true;
      heading.font.size = 12;
      heading.spaceAfter = 6;

      // Insert each bibliography entry as its own paragraph
      for (const entry of entries) {
        const para = body.insertParagraph(entry, Word.InsertLocation.end);
        para.styleBuiltIn = Word.BuiltInStyleName.normal;
        para.font.bold = false;
        para.font.size = 12;
        para.spaceAfter = 4;
      }

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
  // Track all references that have been cited in the document (preserves insertion order)
  private citedReferences: Map<string, Reference> = new Map();
  // Map reference ID → IEEE citation number (1-based, in order of first citation)
  private ieeeNumbers: Map<string, number> = new Map();
  // Cache loaded references so Insert Citation can access them
  private loadedReferences: Reference[] = [];
  // Dynamic citation styles — loaded from backend, fallback to FALLBACK_STYLES
  private availableStyles: string[] = [...FALLBACK_STYLES];

  async initialize(): Promise<void> {
    await Office.onReady();
    const authState = this.auth.getState();
    if (authState.isAuthenticated) {
      this.showProjectView();
    } else {
      // Check if user has seen the welcome screen before
      const hasSeenWelcome = sessionStorage.getItem("lably_has_seen_welcome");
      if (hasSeenWelcome) {
        this.showLoginView();
      } else {
        this.showWelcomeView();
      }
    }
  }

  // ========================================================================
  // WELCOME / FIRST-RUN VIEW
  // ========================================================================

  private showWelcomeView(): void {
    const root = document.getElementById("root")!;
    root.innerHTML = `
      <div class="welcome-container">
        <div class="welcome-icon">
          <img src="assets/icon-80.png" alt="Lably" width="64" height="64" />
        </div>
        <h2>Welcome to Lably</h2>
        <p class="welcome-tagline">Your citation manager, right inside Word.</p>
        <div class="welcome-features">
          <div class="welcome-feature">
            <span class="feature-icon">1</span>
            <p>Search and select references from your Lably projects</p>
          </div>
          <div class="welcome-feature">
            <span class="feature-icon">2</span>
            <p>Insert perfectly formatted inline citations at your cursor</p>
          </div>
          <div class="welcome-feature">
            <span class="feature-icon">3</span>
            <p>Auto-generate and update your bibliography as you write</p>
          </div>
        </div>
        <p class="welcome-styles">Supports APA, MLA, Chicago, Harvard, IEEE and more.</p>
        <div class="welcome-auth-notice">
          <p>Sign in with your Lably account to get started. Don't have an account? You can sign up for free at <a href="https://lably.cloud" target="_blank">lably.cloud</a>.</p>
        </div>
        <button id="getStartedBtn" class="btn btn-primary btn-large">Get Started</button>
        <div class="welcome-links">
          <a href="https://lably.cloud/support" target="_blank">Support</a>
          <span class="link-divider">|</span>
          <a href="https://lably.cloud/privacy" target="_blank">Privacy Policy</a>
        </div>
      </div>
    `;
    document.getElementById("getStartedBtn")!.addEventListener("click", () => {
      sessionStorage.setItem("lably_has_seen_welcome", "true");
      this.showLoginView();
    });
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
        <div class="login-links">
          <p>Don't have an account? <a href="https://lably.cloud/signup" target="_blank">Sign up</a></p>
          <p><a href="https://lably.cloud/support" target="_blank">Need help?</a></p>
        </div>
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
    // Reset cited references when going back to projects
    this.citedReferences.clear();
    this.ieeeNumbers.clear();

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
      // Load projects and available styles in parallel
      const [rawResponse, backendStyles] = await Promise.all([
        this.api.getProjects(),
        this.api.getStyles(),
      ]);
      const projects: Project[] = extractArray(rawResponse);

      // Build the dynamic styles list:
      //  1. If backend provides a styles endpoint, use those
      //  2. Otherwise, merge FALLBACK_STYLES with unique styles found on projects
      if (backendStyles && backendStyles.length > 0) {
        this.availableStyles = backendStyles;
      } else {
        // Start with fallback styles, then add any project-level styles not already present
        const styleSet = new Set(FALLBACK_STYLES.map(s => s.toLowerCase()));
        const merged = [...FALLBACK_STYLES];
        for (const p of projects) {
          const ps = (p.citation_style || "").trim();
          if (ps && !styleSet.has(ps.toLowerCase())) {
            styleSet.add(ps.toLowerCase());
            merged.push(ps);
          }
        }
        this.availableStyles = merged;
      }

      const projectsList = document.getElementById("projectsList")!;
      const loadingDiv = document.getElementById("loadingProjects")!;
      loadingDiv.style.display = "none";

      if (projects.length === 0) {
        projectsList.innerHTML = `<div class="empty-state"><p>No projects found. Create a project in Lably to get started.</p></div>`;
        return;
      }

      projectsList.innerHTML = projects
        .map((project) => `
          <div class="project-card" data-project-id="${escapeHtml(String(project.id || ""))}">
            <h3>${escapeHtml(String(project.name || "Untitled"))}</h3>
            <p class="citation-style">Style: ${escapeHtml(String(project.citation_style || "APA"))}</p>
            <button class="btn btn-primary" data-action="select" data-project-id="${escapeHtml(String(project.id || ""))}">
              Select Project
            </button>
          </div>
        `).join("");

      document.querySelectorAll<HTMLElement>('[data-action="select"]').forEach((btn) => {
        btn.addEventListener("click", () => {
          const projectId = btn.getAttribute("data-project-id")!;
          const project = projects.find((p) => String(p.id) === projectId)!;
          if (project) this.showReferencesView(project);
        });
      });
    } catch (error) {
      const loadingDiv = document.getElementById("loadingProjects")!;
      loadingDiv.innerHTML = `<div class="error-message"><strong>Error:</strong> ${escapeHtml(error instanceof Error ? error.message : String(error))}</div>`;
      loadingDiv.style.display = "block";
    }
  }

  // ========================================================================
  // REFERENCES VIEW
  // ========================================================================

  private async showReferencesView(project: Project): Promise<void> {
    this.currentProject = project;
    this.selectedReferences.clear();
    // Match project style case-insensitively to the available styles list
    const projectStyle = project.citation_style || "";
    this.currentStyle = this.availableStyles.find(s => s.toLowerCase() === projectStyle.toLowerCase()) || this.availableStyles[0] || "APA";

    const root = document.getElementById("root")!;
    root.innerHTML = `
      <div class="references-container">
        <div class="header">
          <button id="backBtn" class="btn btn-secondary btn-small">Back</button>
          <h2>${escapeHtml(String(project.name || ""))}</h2>
        </div>
        <div class="controls">
          <input type="text" id="searchInput" class="search-input" placeholder="Search references..." />
          <div class="style-selector">
            <label for="styleSelect">Citation Style:</label>
            <select id="styleSelect" class="style-select">
              ${this.availableStyles.map((s) => `<option value="${s}" ${s === this.currentStyle ? "selected" : ""}>${s}</option>`).join("")}
            </select>
          </div>
        </div>
        <div id="referencesList" class="references-list"></div>
        <div id="loadingReferences" class="loading">Loading references...</div>
        <div id="statusBar" style="display:none;padding:8px 12px;font-size:12px;background:#e8f5e9;color:#2e7d32;border-top:1px solid #c8e6c9;"></div>
        <div class="action-bar">
          <button id="insertInlineBtn" class="btn btn-primary" disabled>Insert Citation</button>
          <button id="insertBibBtn" class="btn btn-secondary" disabled>Insert Bibliography</button>
        </div>
      </div>
    `;

    document.getElementById("backBtn")!.addEventListener("click", () => this.showProjectView());
    document.getElementById("styleSelect")!.addEventListener("change", (e) => {
      this.currentStyle = (e.target as HTMLSelectElement).value;
      this.loadReferences();
    });
    document.getElementById("searchInput")!.addEventListener("input", () => {
      if (this.searchDebounceTimer) clearTimeout(this.searchDebounceTimer);
      this.searchDebounceTimer = setTimeout(() => this.loadReferences(), 300);
    });
    document.getElementById("insertInlineBtn")!.addEventListener("click", () => this.insertSelectedCitations());
    document.getElementById("insertBibBtn")!.addEventListener("click", () => this.insertBibliographyOnly());

    this.loadReferences();
  }

  private showStatus(msg: string): void {
    const bar = document.getElementById("statusBar");
    if (bar) {
      bar.textContent = msg;
      bar.style.display = "block";
      setTimeout(() => { bar.style.display = "none"; }, 3000);
    }
  }

  private async loadReferences(): Promise<void> {
    if (!this.currentProject) return;
    const referencesList = document.getElementById("referencesList");
    const loadingDiv = document.getElementById("loadingReferences");
    if (!referencesList || !loadingDiv) return;

    loadingDiv.style.display = "block";
    loadingDiv.textContent = "Loading references...";
    loadingDiv.className = "loading";

    try {
      const searchTerm = (document.getElementById("searchInput") as HTMLInputElement)?.value || "";
      const rawResponse = await this.api.getReferences(this.currentProject.id, searchTerm, this.currentStyle);
      const references: Reference[] = extractArray(rawResponse);
      this.loadedReferences = references;

      loadingDiv.style.display = "none";

      if (references.length === 0) {
        referencesList.innerHTML = `<div class="empty-state"><p>No references found.</p></div>`;
        return;
      }

      const styleKey = `formatted_${this.currentStyle.toLowerCase()}`;

      referencesList.innerHTML = references
        .map((ref) => `
          <div class="reference-item" data-ref-id="${escapeHtml(String(ref.id || ""))}">
            <input type="checkbox" class="ref-checkbox" data-ref-id="${escapeHtml(String(ref.id || ""))}"
              ${this.selectedReferences.has(String(ref.id)) ? "checked" : ""} />
            <div class="ref-content">
              <h4>${escapeHtml(String(ref.title || "Untitled"))}</h4>
              <p class="ref-authors">${escapeHtml(authorsToString(ref.authors))}</p>
              <p class="ref-year">${ref.year || ""}</p>
              <p class="ref-formatted">${escapeHtml(String(ref[styleKey] || ""))}</p>
            </div>
          </div>
        `).join("");

      document.querySelectorAll<HTMLInputElement>(".ref-checkbox").forEach((checkbox) => {
        checkbox.addEventListener("change", () => {
          const refId = checkbox.getAttribute("data-ref-id")!;
          if (checkbox.checked) { this.selectedReferences.add(refId); }
          else { this.selectedReferences.delete(refId); }
          this.updateActionButtons();
        });
      });

      this.updateActionButtons();
    } catch (error) {
      loadingDiv.innerHTML = `<div class="error-message"><strong>Error:</strong> ${escapeHtml(error instanceof Error ? error.message : String(error))}</div>`;
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

  /**
   * Insert Citation button: inserts inline citation at cursor AND auto-updates bibliography.
   */
  private async insertSelectedCitations(): Promise<void> {
    if (this.selectedReferences.size === 0) return;

    const btn = document.getElementById("insertInlineBtn") as HTMLButtonElement;
    if (btn) { btn.disabled = true; btn.textContent = "Inserting..."; }

    try {
      const refIds = Array.from(this.selectedReferences);

      // ---- Step 1: Get citation text (try format API, then local fallback) ----
      let citationTextToInsert = "";

      try {
        const rawResponse = await this.api.formatCitations(refIds, this.currentStyle);
        let formatted: Record<string, FormattedCitation> = {};
        if (rawResponse && typeof rawResponse === "object") {
          const firstKey = Object.keys(rawResponse)[0];
          if (firstKey && rawResponse[firstKey]?.inline) {
            formatted = rawResponse;
          } else if (rawResponse.data && typeof rawResponse.data === "object") {
            formatted = rawResponse.data;
          } else if (rawResponse.citations && typeof rawResponse.citations === "object") {
            formatted = rawResponse.citations;
          }
        }
        const inlineParts: string[] = [];
        for (const refId of refIds) {
          if (formatted[refId]?.inline) {
            inlineParts.push(formatted[refId].inline);
          }
        }
        if (inlineParts.length > 0) {
          citationTextToInsert = inlineParts.join(" ");
        }
      } catch {
        // Format API unavailable — use local fallback
      }

      // Fallback: try formatted_inline_* field on references, then build locally
      if (!citationTextToInsert) {
        const inlineKey = `formatted_inline_${this.currentStyle.toLowerCase()}`;
        const parts: string[] = [];
        for (const refId of refIds) {
          const ref = this.loadedReferences.find((r) => String(r.id) === refId);
          if (ref) {
            // Assign IEEE number if not already assigned
            if (!this.ieeeNumbers.has(refId)) {
              this.ieeeNumbers.set(refId, this.ieeeNumbers.size + 1);
            }
            // Prefer backend-provided inline citation field
            if (ref[inlineKey]) {
              parts.push(String(ref[inlineKey]));
            } else {
              parts.push(buildInlineCitation(ref, this.currentStyle, this.ieeeNumbers.get(refId)));
            }
          }
        }
        citationTextToInsert = parts.join(" ");
      }

      // ---- Step 2: Insert citation text into the document ----
      let insertSuccess = false;
      if (citationTextToInsert) {
        try {
          await this.word.insertInlineCitation(citationTextToInsert);
          insertSuccess = true;
        } catch (insertErr) {
          const msg = insertErr instanceof Error ? insertErr.message : String(insertErr);
          this.showStatus("Insert failed: " + msg);
        }
      }

      // ---- Step 3: Track cited references ----
      for (const refId of refIds) {
        const ref = this.loadedReferences.find((r) => String(r.id) === refId);
        if (ref) this.citedReferences.set(String(ref.id), ref);
      }

      // ---- Step 4: Auto-update bibliography ----
      if (this.citedReferences.size > 0) {
        let bibEntries: string[] = [];

        // Try backend bibliography API first
        if (this.currentProject) {
          try {
            const allCitedIds = Array.from(this.citedReferences.keys());
            const bibRaw = await this.api.getBibliography(this.currentProject.id, allCitedIds, this.currentStyle);
            const bibArr = extractArray(bibRaw);
            if (bibArr.length > 0) {
              bibEntries = bibArr.map((entry: any) => {
                if (typeof entry === "string") return entry;
                if (entry.bibliography) return String(entry.bibliography);
                if (entry.formatted) return String(entry.formatted);
                if (entry.text) return String(entry.text);
                return String(entry);
              });
            }
          } catch {
            // Backend bibliography API unavailable — fall through to local
          }
        }

        // Fallback: build locally from formatted_* fields or local rules
        if (bibEntries.length === 0) {
          bibEntries = Array.from(this.citedReferences.entries())
            .map(([id, ref]) => getBibEntry(ref, this.currentStyle, this.ieeeNumbers.get(id)));
        }

        await this.word.insertOrUpdateBibliography(bibEntries);
      }

      if (insertSuccess) {
        this.showStatus(`Inserted ${refIds.length} citation(s) and updated bibliography.`);
      }

      this.selectedReferences.clear();
      document.querySelectorAll<HTMLInputElement>(".ref-checkbox").forEach((cb) => { cb.checked = false; });
      this.updateActionButtons();

    } catch (error) {
      this.showStatus(error instanceof Error ? error.message : "Failed to insert citations");
    } finally {
      if (btn) { btn.disabled = this.selectedReferences.size === 0; btn.textContent = "Insert Citation"; }
    }
  }

  /**
   * Insert Bibliography button: manually insert/update bibliography for selected references.
   */
  private async insertBibliographyOnly(): Promise<void> {
    if (!this.currentProject || this.selectedReferences.size === 0) return;

    const btn = document.getElementById("insertBibBtn") as HTMLButtonElement;
    if (btn) { btn.disabled = true; btn.textContent = "Inserting..."; }

    try {
      const refIds = Array.from(this.selectedReferences);

      // Add selected refs to cited references
      for (const refId of refIds) {
        const ref = this.loadedReferences.find((r) => String(r.id) === refId);
        if (ref) this.citedReferences.set(String(ref.id), ref);
      }

      // Assign IEEE numbers for any new references
      for (const refId of refIds) {
        if (!this.ieeeNumbers.has(refId)) {
          this.ieeeNumbers.set(refId, this.ieeeNumbers.size + 1);
        }
      }

      // Try backend bibliography API first
      let bibEntries: string[] = [];
      if (this.currentProject) {
        try {
          const allCitedIds = Array.from(this.citedReferences.keys());
          const bibRaw = await this.api.getBibliography(this.currentProject.id, allCitedIds, this.currentStyle);
          const bibArr = extractArray(bibRaw);
          if (bibArr.length > 0) {
            bibEntries = bibArr.map((entry: any) => {
              if (typeof entry === "string") return entry;
              if (entry.bibliography) return String(entry.bibliography);
              if (entry.formatted) return String(entry.formatted);
              if (entry.text) return String(entry.text);
              return String(entry);
            });
          }
        } catch {
          // Backend bibliography API unavailable — fall through to local
        }
      }

      // Fallback: build locally from formatted_* fields or local rules
      if (bibEntries.length === 0) {
        bibEntries = Array.from(this.citedReferences.entries())
          .map(([id, ref]) => getBibEntry(ref, this.currentStyle, this.ieeeNumbers.get(id)));
      }

      await this.word.insertOrUpdateBibliography(bibEntries);
      this.showStatus(`Bibliography updated with ${bibEntries.length} entries.`);

    } catch (error) {
      alert(error instanceof Error ? error.message : "Failed to insert bibliography");
    } finally {
      if (btn) { btn.disabled = this.selectedReferences.size === 0; btn.textContent = "Insert Bibliography"; }
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
