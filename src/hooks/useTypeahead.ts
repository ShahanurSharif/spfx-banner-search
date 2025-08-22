import { useState, useRef, useCallback, useEffect } from "react";
import type { SuggestionItem } from "../services/SharePointSearchService";

export function useTypeahead(
  fetchFn: (term: string, signal?: AbortSignal) => Promise<SuggestionItem[]>,
  debounceMs = 250
) {
  const [value, setValue] = useState("");
  const [suggestions, setSuggestions] = useState<SuggestionItem[]>([]);
  const [open, setOpen] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const ctrlRef = useRef<AbortController | null>(null);
  const tRef = useRef<number | null>(null);

  const onChange = useCallback((v: string) => setValue(v ?? ""), []);

  useEffect(() => {
    if (tRef.current) clearTimeout(tRef.current);
    if (!value.trim()) {
      setSuggestions([]);
      setOpen(false);
      return;
    }
    tRef.current = window.setTimeout(async () => {
      if (ctrlRef.current) ctrlRef.current.abort();
      const ctrl = new AbortController();
      ctrlRef.current = ctrl;
      setLoading(true);
      setError(null);
      try {
        const items = await fetchFn(value, ctrl.signal);
        setSuggestions(items);
        setOpen(items.length > 0);
      } catch (e: any) {
        if (e?.name !== "AbortError") {
          setError(e?.message || "Typeahead failed");
          setSuggestions([]);
          setOpen(false);
        }
      } finally {
        setLoading(false);
      }
    }, debounceMs);
    return () => {
      if (tRef.current) clearTimeout(tRef.current);
      if (ctrlRef.current) ctrlRef.current.abort();
    };
  }, [value, fetchFn, debounceMs]);

  return { value, onChange, suggestions, open, loading, error, setOpen, setSuggestions };
}
