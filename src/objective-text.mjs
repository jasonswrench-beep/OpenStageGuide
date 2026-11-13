// Saved objective values are opaque until a feature explicitly needs objective text.
export const unsupportedObjective = module => module?.objective !== undefined && typeof module.objective !== 'string';
export const objectiveText = module => typeof module?.objective === 'string' ? module.objective : '';
export const objectiveLabel = module => unsupportedObjective(module) ? 'Unsupported objective value; preserve the guide and repair this field before reviewing it.' : objectiveText(module) || '(blank)';
