/**
 * Always-allowed tools page — view and revoke persistent approval grants.
 *
 * Grants are created from the tool approval dialog ("Always allow") and are
 * consulted by both the direct-JS gate and the Safe-mode mutation gate.
 * Removing a grant restores the per-call prompt for that tool.
 */

import { t } from "../../../language/index.js";
import type { SettingsShellPage } from "../../../ui/settings-shell.js";
import {
  createNavRow,
  createSettingsGroup,
  createSettingToggleRow,
} from "../../../ui/settings-rows.js";
import { showToast } from "../../../ui/toast.js";
import { getSettingsPagesDependencies } from "./dependencies.js";

function buildEmptyState(): HTMLParagraphElement {
  const empty = document.createElement("p");
  empty.className = "pi-set-empty";
  empty.textContent = t("settings.approvals.empty");
  return empty;
}

export function createApprovalsPage(): SettingsShellPage {
  return {
    id: "approvals",
    parentId: "root",
    title: () => t("settings.page.approvals"),
    subtitle: () => t("settings.page.approvals.sub"),
    render: (ctx) => {
      const deps = getSettingsPagesDependencies();
      const group = createSettingsGroup();
      const emptyState = buildEmptyState();
      const tools = deps.getAlwaysAllowedTools?.() ?? [];

      const showEmpty = (): void => {
        group.list.replaceChildren(emptyState);
      };

      if (tools.length === 0) {
        showEmpty();
      } else {
        for (const toolName of tools) {
          const row = createSettingToggleRow({
            label: toolName,
            sublabel: t("settings.approvals.toolSublabel"),
            checked: true,
            onChange: (checked) => {
              if (checked) return;

              const remove = deps.removeAlwaysAllowedTool;
              if (!remove) return;

              void remove(toolName).then(() => {
                row.root.remove();
                if (group.list.childElementCount === 0) {
                  showEmpty();
                }
                showToast(
                  t("settings.toast.alwaysAllowedRemoved", { tool: toolName }),
                );
              });
            },
          });
          group.list.appendChild(row.root);
        }

        const clearRow = createNavRow({
          label: t("settings.approvals.clearAll"),
          onActivate: () => {
            const clear = deps.clearAlwaysAllowedTools;
            if (!clear) return;

            void clear().then(() => {
              showEmpty();
              showToast(t("settings.toast.alwaysAllowedCleared"));
            });
          },
        });
        group.list.appendChild(clearRow.root);
      }

      ctx.body.appendChild(group.root);
    },
  };
}
