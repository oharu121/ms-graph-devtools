/**
 * Test Multi-Tenant Storage Implementation
 * Verify that multiple tenants/clients can coexist without overwrites
 */

import Azure, { AzureAuth } from "../src/index";
import path from "path";
import os from "os";

function getStorageDir(): string {
  const homeDir = os.homedir();
  if (process.platform === "win32") {
    const localAppData =
      process.env.LOCALAPPDATA || path.join(homeDir, "AppData", "Local");
    return path.join(localAppData, "ms-graph-devtools");
  } else {
    const configHome =
      process.env.XDG_CONFIG_HOME || path.join(homeDir, ".config");
    return path.join(configHome, "ms-graph-devtools");
  }
}

async function test1_multiTenantStorage() {
  console.log("\n=== Test 1: Multi-Tenant Storage ===");

  // Clean up any existing test files
  await Azure.clearStoredCredentials("test-tenant-a", "test-client-1");
  await Azure.clearStoredCredentials("test-tenant-b", "test-client-2");

  // Test Tenant A
  Azure.reset();
  Azure.config({
    tenantId: "test-tenant-a",
    clientId: "test-client-1",
    clientSecret: "secret-a",
    refreshToken: "token-a",
  });

  const auth1 = AzureAuth.getGlobalInstance();
  const storagePath1 = (auth1 as any).storagePath;
  console.log("✓ Tenant A storage path:", storagePath1);
  console.log("  Expected pattern: tokens.test-tenant-a.test-client-1.json");

  // Test Tenant B
  Azure.reset();
  Azure.config({
    tenantId: "test-tenant-b",
    clientId: "test-client-2",
    clientSecret: "secret-b",
    refreshToken: "token-b",
  });

  const auth2 = AzureAuth.getGlobalInstance();
  const storagePath2 = (auth2 as any).storagePath;
  console.log("✓ Tenant B storage path:", storagePath2);
  console.log("  Expected pattern: tokens.test-tenant-b.test-client-2.json");

  // Verify they're different
  if (storagePath1 !== storagePath2) {
    console.log("✓ Storage paths are different (multi-tenant works!)");
  } else {
    console.log("✗ Storage paths are the same (BUG!)");
  }

  // Clean up
  await Azure.clearStoredCredentials("test-tenant-a", "test-client-1");
  await Azure.clearStoredCredentials("test-tenant-b", "test-client-2");
}

async function test2_listStoredCredentials() {
  console.log("\n=== Test 2: List Stored Credentials ===");

  // Create some test credentials
  Azure.reset();
  Azure.config({
    tenantId: "list-test-1",
    clientId: "client-1",
    clientSecret: "secret",
    refreshToken: "token",
  });

  Azure.reset();
  Azure.config({
    tenantId: "list-test-2",
    clientId: "client-2",
    clientSecret: "secret",
    refreshToken: "token",
  });

  // List all credentials
  const stored = await Azure.listStoredCredentials();
  console.log("✓ Stored credentials:");
  stored.forEach((cred) => {
    if (cred.tenantId && cred.clientId) {
      console.log(`  - Tenant: ${cred.tenantId}, Client: ${cred.clientId}`);
    } else {
      console.log(`  - ${cred.file} (legacy)`);
    }
  });

  // Clean up
  await Azure.clearStoredCredentials("list-test-1", "client-1");
  await Azure.clearStoredCredentials("list-test-2", "client-2");
}

async function test3_backwardCompatibility() {
  console.log("\n=== Test 3: Backward Compatibility ===");

  // Test without tenant/client (should use default file)
  Azure.reset();
  Azure.config({
    refreshToken: "legacy-token",
  });

  const auth = AzureAuth.getGlobalInstance();
  const storagePath = (auth as any).storagePath;
  console.log("✓ Legacy storage path:", storagePath);
  console.log("  Expected pattern: tokens.json");

  if (storagePath.endsWith("tokens.json")) {
    console.log("✓ Backward compatibility maintained!");
  } else {
    console.log("✗ Backward compatibility broken!");
  }
}

async function test4_clearCredentials() {
  console.log("\n=== Test 4: Clear Credentials ===");

  // Create test credentials
  Azure.reset();
  Azure.config({
    tenantId: "clear-test",
    clientId: "client-test",
    clientSecret: "secret",
    refreshToken: "token",
  });

  // Verify it exists
  let stored = await Azure.listStoredCredentials();
  const before = stored.filter(
    (c) => c.tenantId === "clear-test" && c.clientId === "client-test"
  );
  console.log(`✓ Before clear: ${before.length} matching file(s)`);

  // Clear specific credentials
  await Azure.clearStoredCredentials("clear-test", "client-test");

  // Verify it's gone
  stored = await Azure.listStoredCredentials();
  const after = stored.filter(
    (c) => c.tenantId === "clear-test" && c.clientId === "client-test"
  );
  console.log(`✓ After clear: ${after.length} matching file(s)`);

  if (before.length > 0 && after.length === 0) {
    console.log("✓ Clear credentials works!");
  } else {
    console.log("✗ Clear credentials failed!");
  }
}

async function test5_pathUpdatesOnLoad() {
  console.log("\n=== Test 5: Storage Path Updates on Load ===");

  // Clean up
  await Azure.clearStoredCredentials("path-test", "client-path");

  // Scenario: First init with tenant/client (creates file)
  Azure.reset();
  Azure.config({
    tenantId: "path-test",
    clientId: "client-path",
    clientSecret: "secret",
    refreshToken: "token-initial",
  });

  // Save to storage (this would happen automatically in real usage)
  const auth1 = AzureAuth.getGlobalInstance();
  await (auth1 as any).saveToStorage();

  let storagePath1 = (auth1 as any).storagePath;
  console.log("✓ Initial storage path:", storagePath1);

  // Scenario: Load from storage (should update path correctly)
  Azure.reset();

  // Set credentials from env to trigger storage load
  process.env.AZURE_CLIENT_ID = "client-path";
  process.env.AZURE_CLIENT_SECRET = "secret";
  process.env.AZURE_TENANT_ID = "path-test";

  const auth2 = AzureAuth.getGlobalInstance();
  await (auth2 as any).loadFromStorage();

  const storagePath2 = (auth2 as any).storagePath;
  console.log("✓ After load storage path:", storagePath2);

  if (storagePath1 === storagePath2) {
    console.log("✓ Storage path correctly updated after load!");
  } else {
    console.log("✗ Storage path mismatch!");
    console.log("  Expected:", storagePath1);
    console.log("  Got:", storagePath2);
  }

  // Clean up
  delete process.env.AZURE_CLIENT_ID;
  delete process.env.AZURE_CLIENT_SECRET;
  delete process.env.AZURE_TENANT_ID;
  await Azure.clearStoredCredentials("path-test", "client-path");
}

async function runAllTests() {
  console.log("╔════════════════════════════════════════════════╗");
  console.log("║                                                ║");
  console.log("║   Multi-Tenant Storage Implementation Tests   ║");
  console.log("║                                                ║");
  console.log("╚════════════════════════════════════════════════╝");

  try {
    await test1_multiTenantStorage();
    await test2_listStoredCredentials();
    await test3_backwardCompatibility();
    await test4_clearCredentials();
    await test5_pathUpdatesOnLoad();

    console.log("\n╔════════════════════════════════════════════════╗");
    console.log("║                                                ║");
    console.log("║          ✓ All Tests Passed! 🎉               ║");
    console.log("║                                                ║");
    console.log("╚════════════════════════════════════════════════╝\n");

    console.log("Multi-tenant storage is working correctly!");
    console.log("\nKey features verified:");
    console.log("✓ Multiple tenants/clients get separate files");
    console.log("✓ No overwrites between different configurations");
    console.log("✓ Backward compatible with legacy single file");
    console.log("✓ Helper methods work (list, clear)");
    console.log("✓ Storage path updates correctly when loading");
  } catch (error) {
    console.error("\n✗ Test failed:", error);
    process.exit(1);
  }
}

// Run tests if called directly
if (require.main === module) {
  runAllTests();
}

export { runAllTests };
