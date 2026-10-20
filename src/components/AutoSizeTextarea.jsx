import React, { useLayoutEffect, useRef } from 'react';

function fitText(field) {
  if (!field || !field.clientWidth) return;
  // Measuring a shorter field can temporarily clamp an ancestor's scroll
  // position. Restore it after measurement so editing long notes stays steady.
  const scrollPositions = [];
  for (let parent = field.parentElement; parent; parent = parent.parentElement) {
    if (parent.scrollTop) scrollPositions.push([parent, parent.scrollTop]);
  }
  const style = getComputedStyle(field);
  const border = parseFloat(style.borderTopWidth) + parseFloat(style.borderBottomWidth);
  field.style.height = 'auto';
  field.style.height = `${Math.ceil(field.scrollHeight + border)}px`;
  field.scrollTop = 0;
  for (const [parent, top] of scrollPositions) parent.scrollTop = top;
}

export function AutoSizeTextarea({ value, className = '', ...props }) {
  const ref = useRef(null);
  useLayoutEffect(() => { fitText(ref.current); }, [value]);
  useLayoutEffect(() => {
    const field = ref.current;
    let width = field.clientWidth;
    const resize = () => {
      if (field.clientWidth === width) return;
      width = field.clientWidth;
      fitText(field);
    };
    const observer = typeof ResizeObserver === 'function' ? new ResizeObserver(resize) : null;
    observer?.observe(field);
    window.addEventListener('resize', resize);
    return () => {
      observer?.disconnect();
      window.removeEventListener('resize', resize);
    };
  }, []);
  return <textarea {...props} ref={ref} value={value} className={`auto-size-textarea ${className}`} />;
}
