import { useState, useRef, useCallback, useEffect } from "react";
import type { SuggestionItem } from "../services/SharePointSearchService";

export function useTypeahead(
  fetchFn: (term: string, signal?: AbortSignal, limit?: number) => Promise<SuggestionItem[]>,
  debounceMs = 250,
  limit = 10,
  zeroTermSuggestions: SuggestionItem[] = []
): {
  value: string;
  onChange: (v: string) => void;
  suggestions: SuggestionItem[];
  open: boolean;
  loading: boolean;
  error: string | undefined;
  setOpen: (open: boolean) => void;
  setSuggestions: (suggestions: SuggestionItem[]) => void;
} {
  const [value, setValue] = useState("");
  const [suggestions, setSuggestions] = useState<SuggestionItem[]>([]);
  const [open, setOpen] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | undefined>(undefined);
  const ctrlRef = useRef<AbortController | null>(null);
  const tRef = useRef<number | null>(null);

  const onChange = useCallback((v: string) => {
    console.debug("[useTypeahead] onChange called with:", v);
    setValue(v ?? "");
  }, []);

  useEffect(() => {
    console.debug("[useTypeahead] useEffect triggered with value:", value);
    if (tRef.current) clearTimeout(tRef.current);
    if (!value.trim()) {
      console.debug("[useTypeahead] Empty value, showing zero-term suggestions");
      setSuggestions(zeroTermSuggestions);
      setOpen(zeroTermSuggestions.length > 0);
      return;
    }
    console.debug("[useTypeahead] Setting timeout for search with value:", value);
    tRef.current = window.setTimeout(async () => {
      if (ctrlRef.current) ctrlRef.current.abort();
      const ctrl = new AbortController();
      ctrlRef.current = ctrl;
      setLoading(true);
      setError(undefined);
      try {
        const items = await fetchFn(value, ctrl.signal, limit);
        console.log("[useTypeahead] Received items count:", items.length);
        console.log("[useTypeahead] Received items:", items);
        setSuggestions(items);
        setOpen(items.length > 0);
        console.log("[useTypeahead] Updated suggestions state, open:", items.length > 0);
      } catch (e: unknown) {
        const error = e as Error;
        if (error?.name !== "AbortError") {
          setError(error?.message || "Typeahead failed");
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
  }, [value, fetchFn, debounceMs, limit, zeroTermSuggestions]);

  return { value, onChange, suggestions, open, loading, error, setOpen, setSuggestions };
}
