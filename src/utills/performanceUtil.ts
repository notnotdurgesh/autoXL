// Performance utility functions for optimizing React component rendering

export interface OptimizedHandlerOptions {
  delay?: number;
  immediate?: boolean;
  throttle?: boolean;
  useRAF?: boolean;
}

/**
 * Creates an optimized handler that debounces or throttles function calls
 * @param handler - The function to optimize
 * @param options - Configuration options for the optimization
 * @returns Optimized handler function
 */
export function createOptimizedHandler(
  handler: (...args: unknown[]) => void,
  options: OptimizedHandlerOptions = {}
): (...args: unknown[]) => void {
  const { delay = 16, immediate = false, throttle = false, useRAF = false } = options;
  let timeoutId: number | null = null;
  let rafId: number | null = null;
  let lastArgs: unknown[];
  let lastCallTime = 0;

  return (...args: unknown[]) => {
    lastArgs = args;
    const now = performance.now();

    const executeHandler = () => {
      handler(...lastArgs);
      lastCallTime = now;
    };

    if (throttle) {
      // Throttling: execute at most once per delay period
      if (now - lastCallTime >= delay) {
        if (useRAF) {
          if (rafId) cancelAnimationFrame(rafId);
          rafId = requestAnimationFrame(executeHandler);
        } else {
          executeHandler();
        }
      }
    } else {
      // Debouncing: delay execution until after delay period of inactivity
      if (immediate && !timeoutId) {
        executeHandler();
      }

      if (timeoutId) {
        clearTimeout(timeoutId);
      }

      if (useRAF) {
        timeoutId = window.setTimeout(() => {
          if (rafId) cancelAnimationFrame(rafId);
          rafId = requestAnimationFrame(() => {
            if (!immediate) {
              executeHandler();
            }
            timeoutId = null;
            rafId = null;
          });
        }, delay);
      } else {
        timeoutId = window.setTimeout(() => {
          if (!immediate) {
            executeHandler();
          }
          timeoutId = null;
        }, delay);
      }
    }
  };
}

/**
 * Runs a function when the browser is idle using requestIdleCallback
 * Falls back to setTimeout if requestIdleCallback is not available
 * @param callback - The function to run when idle
 * @param delay - Optional delay in milliseconds (for setTimeout fallback)
 * @param timeout - Optional timeout in milliseconds (for requestIdleCallback)
 */
export function runWhenIdle(
  callback: () => void, 
  delay: number = 0, 
  timeout: number = 5000
): void {
  if ('requestIdleCallback' in window) {
    requestIdleCallback(callback, { timeout });
  } else {
    // Fallback for browsers that don't support requestIdleCallback
    setTimeout(callback, delay);
  }
}