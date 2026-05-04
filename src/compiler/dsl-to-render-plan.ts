import type { PptDsl } from '../dsl/types';

export interface RenderPlan {
  renderPlanVersion: '1.0';
  renderer: 'wps-com';
  source: {
    dslVersion: string;
    compiledAt: string;
  };
  legacySpec: Record<string, unknown>;
}

export function createRenderPlan(dsl: PptDsl, legacySpec: Record<string, unknown>): RenderPlan {
  return {
    renderPlanVersion: '1.0',
    renderer: 'wps-com',
    source: {
      dslVersion: dsl.dslVersion,
      compiledAt: new Date().toISOString(),
    },
    legacySpec,
  };
}

