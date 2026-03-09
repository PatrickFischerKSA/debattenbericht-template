(function () {
  const defaults = {
    languageToolBaseUrl: "http://localhost:8081/v2/check",
    languageToolLanguage: "de-DE",
    languageToolTimeoutMs: 8000
  };

  window.DEBATTENBERICHT_CONFIG = {
    ...defaults,
    ...(window.DEBATTENBERICHT_CONFIG || {})
  };
})();
