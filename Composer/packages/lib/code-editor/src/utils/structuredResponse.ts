// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { LgTemplate } from '@bfc/shared';

import { emptyTemplateBodyRegex, activityTemplateType } from '../components/lg/constants';
import {
  acceptedAttachmentLayout,
  acceptedInputHintValues,
  AttachmentLayoutStructuredResponse,
  AttachmentsStructuredResponse,
  InputHintStructuredResponse,
  SpeakStructuredResponse,
  StructuredResponse,
  structuredResponseKeys,
  SuggestedActionsStructuredResponse,
  TextStructuredResponse,
} from '../components/lg/types';

const subTemplateNameRegex = /\${(.*)}/;

const getStructuredResponseHelper = (value: unknown, kind: 'Text' | 'Speak' | 'Attachments') => {
  if (typeof value === 'string') {
    const valueAsString = value as string;
    const valueType = subTemplateNameRegex.test(valueAsString) ? 'template' : 'direct';

    return { kind, value: [valueAsString], valueType };
  }

  if (Array.isArray(value) && kind === 'Attachments') {
    const valueAsArray = value as string[];

    return { kind, value: valueAsArray, valueType: 'direct' };
  }

  return undefined;
};

const getStructuredResponseByKind = (
  template: LgTemplate,
  kind: StructuredResponse['kind']
): StructuredResponse | undefined => {
  const value = template.properties?.[kind];
  if (value === undefined) {
    return undefined;
  }

  switch (kind) {
    case 'Text':
      return getStructuredResponseHelper(value, 'Text') as TextStructuredResponse;
    case 'Speak':
      return getStructuredResponseHelper(value, 'Speak') as SpeakStructuredResponse;
    case 'Attachments':
      return getStructuredResponseHelper(value, 'Attachments') as AttachmentsStructuredResponse;
    case 'SuggestedActions': {
      if (Array.isArray(value)) {
        const responseValue = value as string[];
        return { kind: 'SuggestedActions', value: responseValue } as SuggestedActionsStructuredResponse;
      }
      break;
    }
    case 'AttachmentLayout':
      if (acceptedAttachmentLayout.includes(value as typeof acceptedAttachmentLayout[number])) {
        return {
          kind: 'AttachmentLayout',
          value: value as typeof acceptedAttachmentLayout[number],
        } as AttachmentLayoutStructuredResponse;
      }
      break;
    case 'InputHint':
      if (acceptedInputHintValues.includes(value as typeof acceptedInputHintValues[number])) {
        return {
          kind: 'InputHint',
          value: value as typeof acceptedInputHintValues[number],
        } as InputHintStructuredResponse;
      }
      break;
  }

  return undefined;
};

/**
 * Converts template properties to structured response.
 * @param lgTemplate LgTemplate to convert.
 */
export const getStructuredResponseFromTemplate = (
  lgTemplate?: LgTemplate
): Partial<Record<StructuredResponse['kind'], unknown>> | undefined => {
  if (!lgTemplate) {
    return undefined;
  }
  if (!lgTemplate.body || emptyTemplateBodyRegex.test(lgTemplate.body)) {
    return undefined;
  }

  if (lgTemplate.properties?.$type !== activityTemplateType) {
    return undefined;
  }

  const structuredResponse = structuredResponseKeys.reduce((response, key) => {
    const value = getStructuredResponseByKind(lgTemplate, key);
    if (value !== undefined) {
      response[key] = value;
    }

    return response;
  }, {} as Partial<Record<StructuredResponse['kind'], unknown>>);

  return Object.keys(structuredResponse).length ? structuredResponse : undefined;
};

/**
 * Extracts template name from an LG expression
 * ${templateName(params)} => templateName(params)
 * @param expression Expression to extract template name from.
 */
export const extractTemplateNameFromExpression = (expression: string): string | undefined =>
  expression.match(subTemplateNameRegex)?.[1]?.trim();
