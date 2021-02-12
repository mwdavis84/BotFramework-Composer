// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { LgTemplate } from '@botframework-composer/types';

import { BaseEditorProps } from './BaseEditor';
import { LGOption } from './utils/types';

/**
 * Common props for both LG code and response editors.
 */
type LgCommonEditorProps = {
  lgTemplates?: readonly LgTemplate[];
  memoryVariables?: readonly string[];
  lgOption?: LGOption;
  onRemoveTemplate?: (templateId: string) => void;
  onTemplateChange?: (templateId: string, body?: string) => void;
};

/**
 * LG Response editor props;
 */
export type LgResponseEditorProps = LgCommonEditorProps;

/**
 * LG code editor props.
 */
export type LgCodeEditorProps = LgCommonEditorProps &
  BaseEditorProps & {
    toolbarHidden?: boolean;
    onNavigateToLgPage?: (lgFileId: string) => void;
    languageServer?:
      | {
          host?: string;
          hostname?: string;
          port?: number | string;
          path: string;
        }
      | string;
  };
