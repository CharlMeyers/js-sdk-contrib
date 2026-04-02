jest.mock('@azure/app-configuration-provider', () => ({
  load: jest.fn(),
}));

jest.mock('@microsoft/feature-management', () => ({
  ConfigurationMapFeatureFlagProvider: jest.fn(),
  FeatureManager: jest.fn(),
}));

import { AzureFeatureManagerProvider } from './azure-feature-manager-provider';
import { Logger, ErrorCode, ProviderEvents, StandardResolutionReasons } from '@openfeature/web-sdk';
import { ConfigurationMapFeatureFlagProvider, FeatureManager } from '@microsoft/feature-management';
import { load } from '@azure/app-configuration-provider';

describe('AzureFeatureManagerProvider', () => {
  let provider: AzureFeatureManagerProvider;
  let mockAzureAppConfig: any;
  let mockFeatureManager: jest.Mocked<FeatureManager>;
  let mockLogger: jest.Mocked<Logger>;
  const connectionString = 'test-connection-string';
  const refreshInterval = 30000;

  beforeEach(() => {
    jest.useFakeTimers();

    mockAzureAppConfig = {
      refresh: jest.fn(),
    };

    mockFeatureManager = {
      listFeatureNames: jest.fn(),
      isEnabled: jest.fn(),
      getVariant: jest.fn(),
    } as any;

    mockLogger = {
      error: jest.fn(),
    } as any;

    (load as jest.Mock).mockResolvedValue(mockAzureAppConfig);
    (FeatureManager as jest.Mock).mockReturnValue(mockFeatureManager);

    provider = new AzureFeatureManagerProvider(connectionString, refreshInterval, mockLogger);
  });

  afterEach(() => {
    jest.clearAllTimers();
    jest.useRealTimers();
  });

  describe('initialize', () => {
    it('should load Azure App Configuration', async () => {
      mockFeatureManager.listFeatureNames.mockResolvedValue([]);

      await provider.initialize();

      expect(load).toHaveBeenCalledWith(connectionString, {
        featureFlagOptions: {
          enabled: true,
          selectors: [{ keyFilter: '*' }],
          refresh: {
            enabled: true,
            refreshIntervalInMs: refreshInterval,
          },
        },
      });
    });

    it('should create FeatureManager with ConfigurationMapFeatureFlagProvider', async () => {
      mockFeatureManager.listFeatureNames.mockResolvedValue([]);

      await provider.initialize();

      expect(FeatureManager).toHaveBeenCalled();
      expect(ConfigurationMapFeatureFlagProvider).toHaveBeenCalledWith(mockAzureAppConfig);
    });

    it('should prefetch feature flags', async () => {
      const featureNames = ['feature1', 'feature2'];
      mockFeatureManager.listFeatureNames.mockResolvedValue(featureNames);
      mockFeatureManager.isEnabled.mockResolvedValue(true);
      mockFeatureManager.getVariant.mockResolvedValue({ name: 'variant1', configuration: {} });

      await provider.initialize();

      expect(mockFeatureManager.listFeatureNames).toHaveBeenCalled();
      expect(mockFeatureManager.isEnabled).toHaveBeenCalledTimes(2);
    });

    it('should emit Ready event after initialization', async () => {
      mockFeatureManager.listFeatureNames.mockResolvedValue([]);
      const eventSpy = jest.fn();
      provider.events.addHandler(ProviderEvents.Ready, eventSpy);

      await provider.initialize();

      expect(eventSpy).toHaveBeenCalled();
    });

    it('should set up refresh interval', async () => {
      mockFeatureManager.listFeatureNames.mockResolvedValue([]);
      const setIntervalSpy = jest.spyOn(global, 'setInterval');

      await provider.initialize();

      expect(setIntervalSpy).toHaveBeenCalledWith(expect.any(Function), refreshInterval);
      setIntervalSpy.mockRestore();
    });

    it('should refresh configuration on interval', async () => {
      mockFeatureManager.listFeatureNames.mockResolvedValue([]);

      await provider.initialize();

      jest.advanceTimersByTime(refreshInterval);
      await Promise.resolve();

      expect(mockAzureAppConfig.refresh).toHaveBeenCalled();
    });
  });

  describe('prefetchFlags', () => {
    it('should cache enabled feature flags', async () => {
      const featureName = 'enabled-feature';
      mockFeatureManager.listFeatureNames.mockResolvedValue([featureName]);
      mockFeatureManager.isEnabled.mockResolvedValue(true);
      mockFeatureManager.getVariant.mockResolvedValue({ name: 'variant1', configuration: {} });

      await provider.initialize();

      const result = provider.resolveBooleanEvaluation(featureName, false);
      expect(result.value).toBe(true);
    });

    it('should cache disabled feature flags', async () => {
      const featureName = 'disabled-feature';
      mockFeatureManager.listFeatureNames.mockResolvedValue([featureName]);
      mockFeatureManager.isEnabled.mockResolvedValue(false);
      mockFeatureManager.getVariant.mockResolvedValue(undefined);

      await provider.initialize();

      const result = provider.resolveBooleanEvaluation(featureName, true);
      expect(result.value).toBe(false);
    });

    it('should use variant name when available', async () => {
      const featureName = 'feature-with-variant';
      mockFeatureManager.listFeatureNames.mockResolvedValue([featureName]);
      mockFeatureManager.isEnabled.mockResolvedValue(true);
      mockFeatureManager.getVariant.mockResolvedValue({ name: 'custom-variant', configuration: {} });

      await provider.initialize();

      const result = provider.resolveBooleanEvaluation(featureName, false);
      expect(result.variant).toBe('custom-variant');
    });

    it('should use enabled as variant when no variant available', async () => {
      const featureName = 'enabled-without-variant';
      mockFeatureManager.listFeatureNames.mockResolvedValue([featureName]);
      mockFeatureManager.isEnabled.mockResolvedValue(true);
      mockFeatureManager.getVariant.mockResolvedValue(undefined);

      await provider.initialize();

      const result = provider.resolveBooleanEvaluation(featureName, false);
      expect(result.variant).toBe('enabled');
    });

    it('should use disabled as variant when feature is disabled', async () => {
      const featureName = 'disabled-feature';
      mockFeatureManager.listFeatureNames.mockResolvedValue([featureName]);
      mockFeatureManager.isEnabled.mockResolvedValue(false);
      mockFeatureManager.getVariant.mockResolvedValue(undefined);

      await provider.initialize();

      const result = provider.resolveBooleanEvaluation(featureName, false);
      expect(result.variant).toBe('disabled');
    });

    it('should emit ConfigurationChanged event for each feature', async () => {
      const featureNames = ['feature1', 'feature2'];
      mockFeatureManager.listFeatureNames.mockResolvedValue(featureNames);
      mockFeatureManager.isEnabled.mockResolvedValue(true);
      mockFeatureManager.getVariant.mockResolvedValue(undefined);

      const eventSpy = jest.fn();
      provider.events.addHandler(ProviderEvents.ConfigurationChanged, eventSpy);

      await provider.initialize();

      expect(eventSpy).toHaveBeenCalledTimes(2);
    });

    it('should handle errors when fetching individual features', async () => {
      const featureNames = ['good-feature', 'bad-feature'];
      mockFeatureManager.listFeatureNames.mockResolvedValue(featureNames);
      mockFeatureManager.isEnabled.mockImplementation((name) => {
        if (name === 'bad-feature') {
          throw new Error('Feature error');
        }
        return Promise.resolve(true);
      });
      mockFeatureManager.getVariant.mockResolvedValue(undefined);

      await provider.initialize();

      const badResult = provider.resolveBooleanEvaluation('bad-feature', true);
      expect(badResult.value).toBe(false);
      expect(badResult.reason).toBe(StandardResolutionReasons.ERROR);
      expect(badResult.errorCode).toBe(ErrorCode.GENERAL);
    });

    it('should handle errors when listing feature names', async () => {
      mockFeatureManager.listFeatureNames.mockRejectedValue(new Error('List error'));

      await provider.initialize();

      expect(mockLogger.error).toHaveBeenCalledWith(
        'Error populating feature flag cache',
        expect.objectContaining({ error: expect.any(String) }),
      );
    });
  });

  describe('resolveBooleanEvaluation', () => {
    it('should return cached value when available', async () => {
      const featureName = 'cached-feature';
      mockFeatureManager.listFeatureNames.mockResolvedValue([featureName]);
      mockFeatureManager.isEnabled.mockResolvedValue(true);
      mockFeatureManager.getVariant.mockResolvedValue(undefined);

      await provider.initialize();

      const result = provider.resolveBooleanEvaluation(featureName, false);

      expect(result.value).toBe(true);
      expect(result.reason).toBe(StandardResolutionReasons.TARGETING_MATCH);
    });

    it('should return default value for uncached flag', () => {
      const result = provider.resolveBooleanEvaluation('unknown-flag', true);

      expect(result.value).toBe(true);
      expect(result.reason).toBe(StandardResolutionReasons.ERROR);
      expect(result.errorCode).toBe(ErrorCode.FLAG_NOT_FOUND);
    });

    it('should return error when provider not initialized', () => {
      const result = provider.resolveBooleanEvaluation('test-flag', false);

      expect(result.reason).toBe(StandardResolutionReasons.ERROR);
      expect(result.errorMessage).toContain('not initialized');
    });

    it('should use TARGETING_MATCH reason for enabled flags', async () => {
      const featureName = 'enabled-flag';
      mockFeatureManager.listFeatureNames.mockResolvedValue([featureName]);
      mockFeatureManager.isEnabled.mockResolvedValue(true);
      mockFeatureManager.getVariant.mockResolvedValue(undefined);

      await provider.initialize();

      const result = provider.resolveBooleanEvaluation(featureName, false);

      expect(result.reason).toBe(StandardResolutionReasons.TARGETING_MATCH);
    });

    it('should use DEFAULT reason for disabled flags', async () => {
      const featureName = 'disabled-flag';
      mockFeatureManager.listFeatureNames.mockResolvedValue([featureName]);
      mockFeatureManager.isEnabled.mockResolvedValue(false);
      mockFeatureManager.getVariant.mockResolvedValue(undefined);

      await provider.initialize();

      const result = provider.resolveBooleanEvaluation(featureName, true);

      expect(result.reason).toBe(StandardResolutionReasons.DEFAULT);
    });
  });

  describe('resolveStringEvaluation', () => {
    it('should return default value with TYPE_MISMATCH error', () => {
      const result = provider.resolveStringEvaluation('test-flag', 'default');

      expect(result.value).toBe('default');
      expect(result.reason).toBe(StandardResolutionReasons.ERROR);
      expect(result.errorCode).toBe(ErrorCode.TYPE_MISMATCH);
      expect(result.errorMessage).toContain('only supports boolean flags');
    });
  });

  describe('resolveNumberEvaluation', () => {
    beforeEach(() => {
      provider = new AzureFeatureManagerProvider(connectionString, 30000, mockLogger);
    });

    it('should return default value with TYPE_MISMATCH error', () => {
      const result = provider.resolveNumberEvaluation('test-flag', 42);

      expect(result.value).toBe(42);
      expect(result.reason).toBe(StandardResolutionReasons.ERROR);
      expect(result.errorCode).toBe(ErrorCode.TYPE_MISMATCH);
      expect(result.errorMessage).toContain('only supports boolean flags');
    });
  });

  describe('resolveObjectEvaluation', () => {
    beforeEach(() => {
      provider = new AzureFeatureManagerProvider(connectionString, 30000, mockLogger);
    });

    it('should return default value with TYPE_MISMATCH error', () => {
      const defaultObj = { key: 'value' };
      const result = provider.resolveObjectEvaluation('test-flag', defaultObj);

      expect(result.value).toBe(defaultObj);
      expect(result.reason).toBe(StandardResolutionReasons.ERROR);
      expect(result.errorCode).toBe(ErrorCode.TYPE_MISMATCH);
      expect(result.errorMessage).toContain('only supports boolean flags');
    });
  });

  describe('onClose', () => {
    it('should clear refresh interval', async () => {
      mockFeatureManager.listFeatureNames.mockResolvedValue([]);
      const clearIntervalSpy = jest.spyOn(global, 'clearInterval');
      await provider.initialize();

      await provider.onClose();

      expect(clearIntervalSpy).toHaveBeenCalled();
      clearIntervalSpy.mockRestore();
    });

    it('should clear cache', async () => {
      const featureName = 'cached-feature';
      mockFeatureManager.listFeatureNames.mockResolvedValue([featureName]);
      mockFeatureManager.isEnabled.mockResolvedValue(true);
      mockFeatureManager.getVariant.mockResolvedValue(undefined);

      await provider.initialize();
      await provider.onClose();

      const result = provider.resolveBooleanEvaluation(featureName, false);
      expect(result.errorCode).toBe(ErrorCode.FLAG_NOT_FOUND);
    });

    it('should emit Error event', async () => {
      mockFeatureManager.listFeatureNames.mockResolvedValue([]);
      const eventSpy = jest.fn();
      provider.events.addHandler(ProviderEvents.Error, eventSpy);

      await provider.initialize();
      await provider.onClose();

      expect(eventSpy).toHaveBeenCalled();
    });

    it('should handle onClose when not initialized', async () => {
      await expect(provider.onClose()).resolves.not.toThrow();
    });

    it('should handle onClose multiple times', async () => {
      mockFeatureManager.listFeatureNames.mockResolvedValue([]);
      const clearIntervalSpy = jest.spyOn(global, 'clearInterval');
      await provider.initialize();

      await provider.onClose();
      await provider.onClose();
      await provider.onClose();

      expect(clearIntervalSpy).toHaveBeenCalledTimes(1);
      clearIntervalSpy.mockRestore();
    });
  });

  describe('integration scenarios', () => {
    it('should handle full lifecycle', async () => {
      const featureName = 'lifecycle-feature';
      mockFeatureManager.listFeatureNames.mockResolvedValue([featureName]);
      mockFeatureManager.isEnabled.mockResolvedValue(true);
      mockFeatureManager.getVariant.mockResolvedValue(undefined);

      await provider.initialize();
      const result = provider.resolveBooleanEvaluation(featureName, false);
      expect(result.value).toBe(true);

      await provider.onClose();
      const resultAfterClose = provider.resolveBooleanEvaluation(featureName, false);
      expect(resultAfterClose.errorCode).toBe(ErrorCode.FLAG_NOT_FOUND);
    });

    it('should handle multiple features with mixed states', async () => {
      const features = ['enabled1', 'disabled1', 'enabled2'];
      mockFeatureManager.listFeatureNames.mockResolvedValue(features);
      mockFeatureManager.isEnabled.mockImplementation((name) => Promise.resolve(name.startsWith('enabled')));
      mockFeatureManager.getVariant.mockResolvedValue(undefined);

      await provider.initialize();

      expect(provider.resolveBooleanEvaluation('enabled1', false).value).toBe(true);
      expect(provider.resolveBooleanEvaluation('disabled1', true).value).toBe(false);
      expect(provider.resolveBooleanEvaluation('enabled2', false).value).toBe(true);
    });

    it('should update cache on refresh interval', async () => {
      const featureName = 'refresh-feature';
      let callCount = 0;
      mockFeatureManager.listFeatureNames.mockResolvedValue([featureName]);
      mockFeatureManager.isEnabled.mockImplementation(() => {
        callCount++;
        return Promise.resolve(callCount > 1);
      });
      mockFeatureManager.getVariant.mockResolvedValue(undefined);

      await provider.initialize();

      const initialResult = provider.resolveBooleanEvaluation(featureName, true);
      expect(initialResult.value).toBe(false);

      await jest.advanceTimersByTimeAsync(refreshInterval);

      const updatedResult = provider.resolveBooleanEvaluation(featureName, true);
      expect(updatedResult.value).toBe(true);
    });
  });

  describe('edge cases', () => {
    it('should handle empty feature list', async () => {
      mockFeatureManager.listFeatureNames.mockResolvedValue([]);

      await provider.initialize();

      const result = provider.resolveBooleanEvaluation('any-flag', false);
      expect(result.errorCode).toBe(ErrorCode.FLAG_NOT_FOUND);
    });

    it('should handle variant object without name property', async () => {
      const featureName = 'no-variant-name';
      mockFeatureManager.listFeatureNames.mockResolvedValue([featureName]);
      mockFeatureManager.isEnabled.mockResolvedValue(true);
      mockFeatureManager.getVariant.mockResolvedValue({} as any);

      await provider.initialize();

      const result = provider.resolveBooleanEvaluation(featureName, false);
      expect(result.variant).toBe('enabled');
    });

    it('should handle non-Error exceptions in feature fetching', async () => {
      const featureName = 'exception-feature';
      mockFeatureManager.listFeatureNames.mockResolvedValue([featureName]);
      mockFeatureManager.isEnabled.mockRejectedValue('string error');
      mockFeatureManager.getVariant.mockResolvedValue(undefined);

      await provider.initialize();

      const result = provider.resolveBooleanEvaluation(featureName, true);
      expect(result.errorMessage).toContain('Unknown error');
    });
  });
});
