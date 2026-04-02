import type { JsonValue, Logger, Provider, ResolutionDetails } from '@openfeature/web-sdk';
import { ErrorCode, OpenFeatureEventEmitter, ProviderEvents, StandardResolutionReasons } from '@openfeature/web-sdk';
import { ConfigurationMapFeatureFlagProvider, FeatureManager } from '@microsoft/feature-management';
import { load } from '@azure/app-configuration-provider';

export class AzureFeatureManagerProvider implements Provider {
  metadata = {
    name: 'Azure Feature Manager Provider',
  };

  events = new OpenFeatureEventEmitter();

  #refreshInterval: number;
  #refreshTimer: ReturnType<typeof setInterval> | null = null;
  #connectionString: string;
  #featureManager: FeatureManager | null = null;
  #cache: Map<string, ResolutionDetails<boolean>> = new Map();
  #logger?: Logger;

  constructor(connectionString: string, refreshIntervalMs = 30000, logger: Logger) {
    this.#refreshInterval = refreshIntervalMs;
    this.#connectionString = connectionString;
    this.#logger = logger;
  }

  async initialize(): Promise<void> {
    try {
      const azureAppConfig = await load(this.#connectionString, {
        featureFlagOptions: {
          enabled: true,
          selectors: [{ keyFilter: '*' }],
          refresh: {
            enabled: true,
            refreshIntervalInMs: this.#refreshInterval,
          },
        },
      });

      this.#featureManager = new FeatureManager(new ConfigurationMapFeatureFlagProvider(azureAppConfig));
      await this.prefetchFlags();

      this.#refreshTimer = setInterval(async () => {
        azureAppConfig.refresh();
        await this.prefetchFlags();
      }, this.#refreshInterval);

      this.events.emit(ProviderEvents.Ready);
    } catch (err) {
      this.#logger?.error('Error initializing Azure Feature Manager Provider', {
        error: err instanceof Error ? err.message : String(err),
      });
      this.events.emit(ProviderEvents.Error);
    }
  }

  async onClose(): Promise<void> {
    if (this.#refreshTimer) {
      clearInterval(this.#refreshTimer as any);
      this.#refreshTimer = null;
    }

    this.#cache.clear();
    this.#featureManager = null;
    this.events.emit(ProviderEvents.Error);
  }

  resolveBooleanEvaluation(flagKey: string, defaultValue: boolean): ResolutionDetails<boolean> {
    const cached = this.#cache.get(flagKey) as ResolutionDetails<boolean> | undefined;
    if (cached && typeof cached.value === 'boolean') {
      return cached;
    }

    return {
      value: defaultValue,
      reason: StandardResolutionReasons.ERROR,
      errorCode: !cached ? ErrorCode.FLAG_NOT_FOUND : ErrorCode.PROVIDER_NOT_READY,
      errorMessage: 'Feature value not available (provider not initialized or cache miss).',
    };
  }

  resolveStringEvaluation(flagKey: string, defaultValue: string): ResolutionDetails<string> {
    return {
      value: defaultValue,
      reason: StandardResolutionReasons.ERROR,
      errorCode: ErrorCode.TYPE_MISMATCH,
      errorMessage: 'Feature Manager only supports boolean flags',
    };
  }

  resolveNumberEvaluation(flagKey: string, defaultValue: number): ResolutionDetails<number> {
    return {
      value: defaultValue,
      reason: StandardResolutionReasons.ERROR,
      errorCode: ErrorCode.TYPE_MISMATCH,
      errorMessage: 'Feature Manager only supports boolean flags',
    };
  }

  resolveObjectEvaluation<T extends JsonValue>(flagKey: string, defaultValue: T): ResolutionDetails<T> {
    return {
      value: defaultValue,
      reason: StandardResolutionReasons.ERROR,
      errorCode: ErrorCode.TYPE_MISMATCH,
      errorMessage: 'Feature Manager only supports boolean flags',
    };
  }

  private async prefetchFlags(): Promise<void> {
    if (!this.#featureManager) return;

    try {
      const names = await this.#featureManager.listFeatureNames();

      for (const name of names) {
        try {
          const enabled = await this.#featureManager.isEnabled(name);
          const variantObj = await this.#featureManager.getVariant(name);
          const variant = (variantObj && variantObj.name) || (enabled ? 'enabled' : 'disabled');

          const details: ResolutionDetails<boolean> = {
            value: enabled,
            reason: enabled ? StandardResolutionReasons.TARGETING_MATCH : StandardResolutionReasons.DEFAULT,
            variant: variant as string,
          };

          this.#cache.set(name, details);
          this.events.emit(ProviderEvents.ConfigurationChanged);
        } catch (innerErr) {
          this.#cache.set(name, {
            value: false,
            reason: StandardResolutionReasons.ERROR,
            errorCode: ErrorCode.GENERAL,
            errorMessage: innerErr instanceof Error ? innerErr.message : 'Unknown error while populating cache',
          });
        }
      }
    } catch (err) {
      this.#logger?.error('Error populating feature flag cache', {
        error: err instanceof Error ? err.message : String(err),
      });
    }
  }
}
