import { AzureFeatureManagerProvider } from './azure-feature-manager-provider';

describe('AzureFeatureManagerProvider', () => {
  it('should be and instance of AzureFeatureManagerProvider', () => {
    expect(new AzureFeatureManagerProvider()).toBeInstanceOf(AzureFeatureManagerProvider);
  });
});
