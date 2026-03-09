(function () {
  class LanguageToolServiceError extends Error {
    constructor(code, message, details = "") {
      super(message);
      this.name = "LanguageToolServiceError";
      this.code = code;
      this.details = details;
    }
  }

  class LanguageToolService {
    constructor(config = {}) {
      this.baseUrl = config.baseUrl || "http://localhost:8081/v2/check";
      this.language = config.language || "de-DE";
      this.timeoutMs = Number(config.timeoutMs || 8000);
    }

    async checkText(text, options = {}) {
      const baseUrl = options.baseUrl || this.baseUrl;
      const language = options.language || this.language;
      const timeoutMs = Number(options.timeoutMs || this.timeoutMs);
      const payload = new URLSearchParams({
        text: String(text || ""),
        language
      });

      const controller = new AbortController();
      const timeoutId = window.setTimeout(() => controller.abort(), timeoutMs);

      let response;
      try {
        response = await fetch(baseUrl, {
          method: "POST",
          headers: {
            "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
            Accept: "application/json"
          },
          body: payload.toString(),
          signal: controller.signal
        });
      } catch (error) {
        window.clearTimeout(timeoutId);
        if (error && error.name === "AbortError") {
          throw new LanguageToolServiceError("timeout", "LanguageTool-Server hat nicht rechtzeitig geantwortet.");
        }
        throw new LanguageToolServiceError(
          "network",
          "LanguageTool-Server nicht erreichbar oder durch CORS blockiert.",
          error && error.message ? error.message : ""
        );
      }

      window.clearTimeout(timeoutId);

      if (!response.ok) {
        throw new LanguageToolServiceError(
          "http",
          `LanguageTool antwortete mit HTTP ${response.status}.`,
          response.statusText || ""
        );
      }

      const raw = await response.text();
      if (!raw || !raw.trim()) {
        throw new LanguageToolServiceError("empty", "LanguageTool lieferte eine leere Antwort.");
      }

      let data;
      try {
        data = JSON.parse(raw);
      } catch {
        throw new LanguageToolServiceError("invalid", "LanguageTool lieferte kein gueltiges JSON.");
      }

      if (!data || !Array.isArray(data.matches)) {
        throw new LanguageToolServiceError("invalid", "LanguageTool-Antwort hat kein gueltiges matches-Array.");
      }

      return data;
    }
  }

  window.LanguageToolService = LanguageToolService;
  window.LanguageToolServiceError = LanguageToolServiceError;
})();
