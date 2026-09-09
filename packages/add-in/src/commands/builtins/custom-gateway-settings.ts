/**
 * Settings section for custom OpenAI-compatible gateways.
 */

import { getAppStorage } from "../../storage/local/app-storage.js";

import {
  DEFAULT_OPENAI_GATEWAY_CONTEXT_WINDOW,
  DEFAULT_GATEWAY_API_TYPE,
  deleteOpenAiGatewayConfig,
  listOpenAiGatewayConfigs,
  normalizeGatewayEndpointUrl,
  saveOpenAiGatewayConfig,
  type OpenAiGatewayConfig,
} from "../../auth/custom-gateways.js";
import type { BrowserProviderApi } from "../../models/browser-model-runtime.js";
import { requestConfirmationDialog } from "../../ui/confirm-dialog.js";
import {
  createButton,
  createConfigInput,
  createConfigRow,
} from "../../ui/extensions-hub-components.js";
import { createOverlaySectionTitle } from "../../ui/overlay-dialog.js";
import { showToast } from "../../ui/toast.js";
import { t } from "../../language/index.js";

interface BuildCustomGatewaySectionOptions {
  onProvidersChanged: () => void;
  /**
   * Include the section's own title + hint. Defaults to true; the settings
   * gateway page supplies its own page header instead.
   */
  includeHeading?: boolean;
}

/**
 * Auto-detect the API type from the endpoint URL path.
 * - `/messages` → anthropic-messages
 * - `/responses` → openai-responses
 * - everything else → openai-completions
 */
function autoDetectApiType(endpointUrl: string): BrowserProviderApi {
  try {
    const path = new URL(endpointUrl).pathname.toLowerCase();
    if (path.includes("/messages")) return "anthropic-messages";
    if (path.includes("/responses")) return "openai-responses";
  } catch {
    // Invalid URL — fall through to default.
  }
  return DEFAULT_GATEWAY_API_TYPE;
}

/**
 * API type → URL path segment for probing.
 * The /models endpoint is at baseUrl/models regardless of api type,
 * but we need the api-specific base to construct probe requests.
 */
const API_TYPE_PROBE_PATHS: Record<BrowserProviderApi, string> = {
  "openai-completions": "/chat/completions",
  "openai-responses": "/responses",
  "anthropic-messages": "/messages",
};

/**
 * Probe the gateway to determine which API type each model supports.
 * Sends minimal POST requests to each API endpoint for a sample of models.
 * Returns a modelId → apiType map.
 */
async function probeGatewayApiTypes(
  endpointUrl: string,
  apiKey: string,
  modelIds: string[],
  fetchFn: typeof globalThis.fetch,
): Promise<Record<string, BrowserProviderApi>> {
  const result: Record<string, BrowserProviderApi> = {};

  // Parse the endpoint URL to get the base path
  let parsed: URL;
  try {
    parsed = new URL(endpointUrl);
  } catch {
    return result;
  }

  // Remove the api-specific path suffix (e.g., /chat/completions) to get the base
  const pathname = parsed.pathname
    .replace(/\/chat\/completions$/i, "")
    .replace(/\/responses$/i, "")
    .replace(/\/messages$/i, "")
    .replace(/\/+$/, "");
  const baseUrl = `${parsed.origin}${pathname}`;

  // Probe the first3 models against all3 API types
  const sampleModels = modelIds.slice(0, 3);
  const apiTypes: BrowserProviderApi[] = [
    "openai-completions",
    "openai-responses",
    "anthropic-messages",
  ];

  for (const modelId of sampleModels) {
    for (const apiType of apiTypes) {
      const probePath = API_TYPE_PROBE_PATHS[apiType];
      const probeUrl = `${baseUrl}${probePath}`;

      try {
        const headers: Record<string, string> = {
          "Content-Type": "application/json",
        };
        if (apiKey) {
          headers["Authorization"] = `Bearer ${apiKey}`;
        }

        // Send a minimal request — just enough to check if the model/API combination works
        const body = JSON.stringify({
          model: modelId,
          messages: [{ role: "user", content: "hi" }],
          max_tokens: 1,
        });

        const response = await fetchFn(probeUrl, {
          method: "POST",
          headers,
          body,
          signal: AbortSignal.timeout(5000),
        });

        // If we get200-299, the model works with this API type
        if (response.status >= 200 && response.status < 300) {
          result[modelId] = apiType;
          break; // Found the right API type for this model
        }

        // If we get401/403, the API type is valid but auth failed — still means this API type works
        if (response.status === 401 || response.status === 403) {
          result[modelId] = apiType;
          break;
        }

        // If we get400 with a model-related error (not endpoint-not-found), the API type is valid
        if (response.status === 400) {
          const text = await response.text().catch(() => "");
          if (text.includes("model") || text.includes("Model")) {
            result[modelId] = apiType;
            break;
          }
        }

        // 404 or other error — wrong API type, try the next one
      } catch {
        // Network error or timeout — try the next API type
        continue;
      }
    }
  }

  return result;
}

function createHint(text: string): HTMLParagraphElement {
  const hint = document.createElement("p");
  hint.className = "pi-overlay-hint";
  hint.textContent = text;
  return hint;
}

function formatTokenCount(value: number): string {
  return `${value.toLocaleString()} tokens`;
}

function createGatewayCard(args: {
  gateway: OpenAiGatewayConfig;
  onEdit: (gateway: OpenAiGatewayConfig) => void;
  onDelete: (gateway: OpenAiGatewayConfig) => void;
}): HTMLElement {
  const card = document.createElement("div");
  card.className = "pi-overlay-surface pi-settings-gateway-item";

  const topRow = document.createElement("div");
  topRow.className = "pi-settings-gateway-item__top";

  const titleGroup = document.createElement("div");
  titleGroup.className = "pi-settings-gateway-item__title-group";

  const title = document.createElement("p");
  title.className = "pi-settings-gateway-item__title";
  title.textContent = args.gateway.displayName;

  const provider = document.createElement("p");
  provider.className = "pi-settings-gateway-item__provider";
  provider.textContent = args.gateway.providerName;

  titleGroup.append(title, provider);

  const actions = document.createElement("div");
  actions.className = "pi-settings-gateway-item__actions";

  const editButton = createButton(t("custom-gateway.editButton"), {
    compact: true,
    onClick: () => {
      args.onEdit(args.gateway);
    },
  });

  const deleteButton = createButton(t("custom-gateway.deleteButton"), {
    compact: true,
    danger: true,
    onClick: () => {
      args.onDelete(args.gateway);
    },
  });

  actions.append(editButton, deleteButton);
  topRow.append(titleGroup, actions);

  const endpoint = document.createElement("p");
  endpoint.className = "pi-settings-gateway-item__meta";
  endpoint.textContent = t("custom-gateway.gatewayEndpoint", {
    url: args.gateway.endpointUrl,
  });

  const model = document.createElement("p");
  model.className = "pi-settings-gateway-item__meta";
  model.textContent = t("custom-gateway.gatewayModel", {
    id: args.gateway.modelId,
  });

  const contextWindow = document.createElement("p");
  contextWindow.className = "pi-settings-gateway-item__meta";
  contextWindow.textContent = t("custom-gateway.gatewayContextWindow", {
    tokens: formatTokenCount(args.gateway.contextWindow),
  });

  const apiType = document.createElement("p");
  apiType.className = "pi-settings-gateway-item__meta";
  apiType.textContent = t("custom-gateway.gatewayApiType", {
    type: args.gateway.apiType,
  });

  const keyState = document.createElement("p");
  keyState.className = "pi-settings-gateway-item__meta";
  keyState.textContent =
    args.gateway.apiKey.length > 0
      ? t("custom-gateway.gatewayApiKeyConfigured")
      : t("custom-gateway.gatewayApiKeyNone");

  card.append(topRow, endpoint, model, contextWindow, apiType, keyState);
  return card;
}

export async function buildCustomGatewaySection(
  options: BuildCustomGatewaySectionOptions,
): Promise<HTMLElement> {
  const section = document.createElement("section");
  section.className = "pi-overlay-section pi-settings-section";
  section.dataset.settingsAnchor = "custom-gateways";

  const title = createOverlaySectionTitle(t("custom-gateway.title"));
  const hint = createHint(t("custom-gateway.hint"));

  const content = document.createElement("div");
  content.className = "pi-settings-section__content";

  const formCard = document.createElement("div");
  formCard.className = "pi-overlay-surface pi-settings-gateway-form";

  const nameInput = createConfigInput({
    placeholder: t("custom-gateway.namePlaceholder"),
  });

  const endpointInput = createConfigInput({
    placeholder: t("custom-gateway.endpointPlaceholder"),
  });
  endpointInput.spellcheck = false;

  const modelInput = createConfigInput({
    placeholder: t("custom-gateway.modelPlaceholder"),
  });

  const contextWindowInput = createConfigInput({
    placeholder: String(DEFAULT_OPENAI_GATEWAY_CONTEXT_WINDOW),
    type: "number",
  });
  contextWindowInput.min = "1024";
  contextWindowInput.step = "1";
  contextWindowInput.inputMode = "numeric";

  const apiKeyInput = createConfigInput({
    placeholder: t("custom-gateway.apiKeyPlaceholder"),
    type: "password",
  });

  const errorText = document.createElement("p");
  errorText.className = "pi-overlay-hint pi-overlay-text-warning";
  errorText.hidden = true;

  const formActions = document.createElement("div");
  formActions.className = "pi-overlay-actions";

  const cancelButton = createButton(t("custom-gateway.cancelButton"), {
    compact: true,
  });
  cancelButton.hidden = true;

  const saveButton = createButton(t("custom-gateway.saveGateway"), {
    compact: true,
    primary: true,
  });

  formActions.append(cancelButton, saveButton);

  formCard.append(
    createConfigRow(t("custom-gateway.configLabelName"), nameInput),
    createConfigRow(t("custom-gateway.configLabelEndpoint"), endpointInput),
    createConfigRow(t("custom-gateway.configLabelModel"), modelInput),
    createConfigRow(
      t("custom-gateway.configLabelContextWindow"),
      contextWindowInput,
    ),
    createHint(t("custom-gateway.contextWindowHint")),
    createConfigRow(t("custom-gateway.configLabelApiKey"), apiKeyInput),
    errorText,
    formActions,
  );

  const listTitle = document.createElement("p");
  listTitle.className = "pi-settings-gateway-list__title";
  listTitle.textContent = t("custom-gateway-settings.configured-gateways");

  const listHost = document.createElement("div");
  listHost.className = "pi-settings-gateway-list";

  let editingGatewayId: string | null = null;
  let gateways: OpenAiGatewayConfig[] = [];

  const setError = (message: string | null): void => {
    if (!message) {
      errorText.hidden = true;
      errorText.textContent = "";
      return;
    }

    errorText.hidden = false;
    errorText.textContent = message;
  };

  const resetForm = (): void => {
    editingGatewayId = null;
    nameInput.value = "";
    endpointInput.value = "";
    modelInput.value = "";
    contextWindowInput.value = "";
    apiKeyInput.value = "";
    cancelButton.hidden = true;
    saveButton.textContent = t("custom-gateway.saveGateway");
    setError(null);
  };

  const startEditing = (gateway: OpenAiGatewayConfig): void => {
    editingGatewayId = gateway.id;
    nameInput.value = gateway.displayName;
    endpointInput.value = gateway.endpointUrl;
    modelInput.value = gateway.modelId;
    contextWindowInput.value = String(gateway.contextWindow);
    apiKeyInput.value = gateway.apiKey;
    cancelButton.hidden = false;
    saveButton.textContent = t("custom-gateway-settings.update-gateway");
    setError(null);
    nameInput.focus();
  };

  const reloadGateways = async (): Promise<void> => {
    gateways = await listOpenAiGatewayConfigs(getAppStorage().customProviders);
  };

  const renderList = (): void => {
    listHost.replaceChildren();

    if (gateways.length === 0) {
      listHost.appendChild(createHint(t("custom-gateway.noGateways")));
      return;
    }

    for (const gateway of gateways) {
      listHost.appendChild(
        createGatewayCard({
          gateway,
          onEdit: startEditing,
          onDelete: (targetGateway) => {
            void (async () => {
              try {
                const confirmed = await requestConfirmationDialog({
                  title: t("custom-gateway.deleteConfirmTitle"),
                  message: t("custom-gateway.deleteConfirmMsg", {
                    name: targetGateway.displayName,
                  }),
                  confirmLabel: t("custom-gateway.deleteButton"),
                  confirmButtonTone: "danger",
                  restoreFocusOnClose: false,
                });

                if (!confirmed) {
                  return;
                }

                await deleteOpenAiGatewayConfig(
                  getAppStorage().customProviders,
                  targetGateway.id,
                );
                await reloadGateways();
                renderList();
                options.onProvidersChanged();
                showToast(
                  t("custom-gateway.deletedGateway", {
                    name: targetGateway.displayName,
                  }),
                );

                if (editingGatewayId === targetGateway.id) {
                  resetForm();
                }
              } catch (error) {
                const message =
                  error instanceof Error ? error.message : String(error);
                showToast(t("custom-gateway.toast.deleteFailed", { message }));
              }
            })();
          },
        }),
      );
    }
  };

  cancelButton.addEventListener("click", () => {
    resetForm();
  });

  saveButton.addEventListener("click", () => {
    void (async () => {
      setError(null);
      saveButton.disabled = true;
      cancelButton.disabled = true;

      try {
        const rawContextWindow = contextWindowInput.value.trim();
        const contextWindow =
          rawContextWindow.length > 0 ? Number(rawContextWindow) : undefined;

        const detectedApiType = autoDetectApiType(endpointInput.value);

        // Probe the gateway to determine per-model API types
        let modelApiMap: Record<string, BrowserProviderApi> | undefined;
        try {
          const endpointUrl = normalizeGatewayEndpointUrl(endpointInput.value);
          // Fetch the /models endpoint to get all model IDs
          const modelsUrl = `${endpointUrl}/models`;
          const modelsResponse = await globalThis.fetch(modelsUrl, {
            signal: AbortSignal.timeout(8000),
          });
          if (modelsResponse.ok) {
            const modelsData = (await modelsResponse
              .json()
              .catch(() => null)) as { data?: { id: string }[] } | null;
            const allModelIds =
              modelsData?.data?.map((m) => m.id).filter(Boolean) ?? [];
            if (allModelIds.length > 0) {
              modelApiMap = await probeGatewayApiTypes(
                endpointUrl,
                apiKeyInput.value,
                allModelIds,
                globalThis.fetch,
              );
            }
          }
        } catch {
          // Probing failed — continue with gateway default API type
        }

        const saved = await saveOpenAiGatewayConfig(
          getAppStorage().customProviders,
          {
            ...(editingGatewayId ? { id: editingGatewayId } : {}),
            displayName: nameInput.value,
            endpointUrl: endpointInput.value,
            modelId: modelInput.value,
            apiKey: apiKeyInput.value,
            ...(contextWindow !== undefined ? { contextWindow } : {}),
            apiType: detectedApiType,
            ...(modelApiMap ? { modelApiMap } : {}),
          },
        );

        await reloadGateways();
        renderList();
        options.onProvidersChanged();
        showToast(
          editingGatewayId
            ? t("custom-gateway.updatedGateway", { name: saved.displayName })
            : t("custom-gateway.savedGateway", { name: saved.displayName }),
        );

        resetForm();
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        setError(message);
      } finally {
        saveButton.disabled = false;
        cancelButton.disabled = false;
      }
    })();
  });

  await reloadGateways();
  renderList();

  content.append(formCard, listTitle, listHost);
  if (options.includeHeading !== false) {
    section.append(title, hint);
  }
  section.append(content);
  return section;
}
