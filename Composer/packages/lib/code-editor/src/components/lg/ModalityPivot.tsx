// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { LgTemplate } from '@bfc/shared';
import { FluentTheme, FontSizes } from '@uifabric/fluent-theme';
import formatMessage from 'format-message';
import { IconButton } from 'office-ui-fabric-react/lib/Button';
import {
  ContextualMenuItemType,
  IContextualMenuItem,
  IContextualMenuItemProps,
  IContextualMenuItemRenderFunctions,
  IContextualMenuProps,
} from 'office-ui-fabric-react/lib/ContextualMenu';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { IPivotStyles, Pivot, PivotItem } from 'office-ui-fabric-react/lib/Pivot';
import { Stack } from 'office-ui-fabric-react/lib/Stack';
import React, { useCallback, useMemo, useRef, useState } from 'react';
import mergeWith from 'lodash/mergeWith';

import { LGOption } from '../../utils';
import { ItemWithTooltip } from '../ItemWithTooltip';

import { AttachmentModalityEditor } from './modalityEditors/AttachmentModalityEditor';
import { SpeechModalityEditor } from './modalityEditors/SpeechModalityEditor';
import { SuggestedActionsModalityEditor } from './modalityEditors/SuggestedActionsModalityEditor';
import { TextModalityEditor } from './modalityEditors/TextModalityEditor';
import {
  AttachmentsStructuredResponse,
  SpeakStructuredResponse,
  SuggestedActionsStructuredResponse,
  TextStructuredResponse,
  ModalityType,
  TemplateResponse,
} from './types';

const modalityDocumentUrl =
  'https://docs.microsoft.com/en-us/azure/bot-service/language-generation/language-generation-structured-response-template?view=azure-bot-service-4.0';

const getModalityTooltipText = (modality: ModalityType) => {
  switch (modality) {
    case 'Attachments':
      return formatMessage(
        'List of attachments with their type. Used by channels to render as UI cards or other generic file attachment types.'
      );
    case 'Speak':
      return formatMessage('Spoken text used by the channel to render audibly.');
    case 'SuggestedActions':
      return formatMessage('List of actions rendered as suggestions to user.');
    case 'Text':
      return formatMessage('Display text used by the channel to render visually.');
  }
};

const addButtonIconProps = { iconName: 'Add', styles: { root: { fontSize: FontSizes.size14 } } };

const styles: { tabs: Partial<IPivotStyles> } = {
  tabs: {
    link: {
      fontSize: FontSizes.size12,
    },
    linkIsSelected: {
      fontSize: FontSizes.size12,
    },
  },
};

const renderModalityEditor = (
  response: TemplateResponse | undefined,
  modality: string,
  onRemoveModality: (modality: ModalityType) => () => void,
  onTemplateChange: (templateId: string, body?: string) => void,
  onAttachmentLayoutChange: (layout: string) => void,
  onInputHintChange: (inputHintString: string) => void,
  onUpdateResponseTemplate: (response: TemplateResponse) => void,
  disableRemoveModality: boolean,
  lgOption?: LGOption,
  lgTemplates?: readonly LgTemplate[],
  memoryVariables?: readonly string[]
) => {
  switch (modality) {
    case 'Attachments':
      return (
        <AttachmentModalityEditor
          lgOption={lgOption}
          lgTemplates={lgTemplates}
          memoryVariables={memoryVariables}
          removeModalityDisabled={disableRemoveModality}
          response={response?.Attachments as AttachmentsStructuredResponse}
          onAttachmentLayoutChange={onAttachmentLayoutChange}
          onRemoveModality={onRemoveModality('Attachments')}
          onTemplateChange={onTemplateChange}
          onUpdateResponseTemplate={onUpdateResponseTemplate}
        />
      );
    case 'Speak':
      return (
        <SpeechModalityEditor
          lgOption={lgOption}
          lgTemplates={lgTemplates}
          memoryVariables={memoryVariables}
          removeModalityDisabled={disableRemoveModality}
          response={response?.Speak as SpeakStructuredResponse}
          onInputHintChange={onInputHintChange}
          onRemoveModality={onRemoveModality('Speak')}
          onTemplateChange={onTemplateChange}
          onUpdateResponseTemplate={onUpdateResponseTemplate}
        />
      );
    case 'SuggestedActions':
      return (
        <SuggestedActionsModalityEditor
          lgOption={lgOption}
          lgTemplates={lgTemplates}
          memoryVariables={memoryVariables}
          removeModalityDisabled={disableRemoveModality}
          response={response?.Speak as SuggestedActionsStructuredResponse}
          onRemoveModality={onRemoveModality('SuggestedActions')}
          onTemplateChange={onTemplateChange}
          onUpdateResponseTemplate={onUpdateResponseTemplate}
        />
      );
    case 'Text':
      return (
        <TextModalityEditor
          lgOption={lgOption}
          lgTemplates={lgTemplates}
          memoryVariables={memoryVariables}
          removeModalityDisabled={disableRemoveModality}
          response={response?.Text as TextStructuredResponse}
          onRemoveModality={onRemoveModality('Text')}
          onTemplateChange={onTemplateChange}
          onUpdateResponseTemplate={onUpdateResponseTemplate}
        />
      );
  }
};

const getInitialModalities = (response?: TemplateResponse): ModalityType[] => {
  const modalities = Object.keys(response || {}) as ModalityType[];
  return modalities.length ? modalities : ['Text'];
};

type Props = {
  lgOption?: LGOption;
  lgTemplates?: readonly LgTemplate[];
  memoryVariables?: readonly string[];
  response?: TemplateResponse;
  onTemplateChange?: (templateId: string, body?: string) => void;
};

const ModalityPivot = React.memo((props: Props) => {
  const { lgOption, lgTemplates, memoryVariables, response: initialResponse, onTemplateChange = () => {} } = props;

  const [response, setResponse] = React.useState(initialResponse);

  const containerRef = useRef<HTMLDivElement>(null);
  const [modalities, setModalities] = useState<ModalityType[]>(getInitialModalities(response));
  const [selectedKey, setSelectedKey] = useState<string>(modalities[0] as string);

  const renderMenuItemContent = React.useCallback(
    (itemProps: IContextualMenuItemProps, defaultRenders: IContextualMenuItemRenderFunctions) =>
      itemProps.item.itemType === ContextualMenuItemType.Header ? (
        <ItemWithTooltip
          itemText={defaultRenders.renderItemName(itemProps)}
          tooltipId="modality-add-menu-header"
          tooltipText={formatMessage.rich('To learn more about modalities, <a>go to this document</a>.', {
            a: ({ children }) => (
              <Link key="modality-add-menu-header-link" href={modalityDocumentUrl} target="_blank">
                {children}
              </Link>
            ),
          })}
        />
      ) : (
        <ItemWithTooltip
          itemText={defaultRenders.renderItemName(itemProps)}
          tooltipId={itemProps.item.key}
          tooltipText={getModalityTooltipText(itemProps.item.key as ModalityType)}
        />
      ),
    []
  );

  const items = useMemo<IContextualMenuItem[]>(
    () => [
      {
        key: 'header',
        itemType: ContextualMenuItemType.Header,
        text: formatMessage('Add modality to this response'),
        onRenderContent: renderMenuItemContent,
      },
      {
        key: 'Text',
        text: formatMessage('Text'),
        onRenderContent: renderMenuItemContent,
        style: { fontSize: FluentTheme.fonts.small.fontSize },
      },
      {
        key: 'Speak',
        text: formatMessage('Speech'),
        onRenderContent: renderMenuItemContent,
        style: { fontSize: FluentTheme.fonts.small.fontSize },
      },
      {
        key: 'Attachments',
        text: formatMessage('Attachments'),
        onRenderContent: renderMenuItemContent,
        style: { fontSize: FluentTheme.fonts.small.fontSize },
      },
      {
        key: 'SuggestedActions',
        text: formatMessage('Suggested Actions'),
        onRenderContent: renderMenuItemContent,
        style: { fontSize: FluentTheme.fonts.small.fontSize },
      },
    ],
    [renderMenuItemContent]
  );

  const pivotItems = useMemo(
    () =>
      modalities.map((modality) => items.find(({ key }) => key === modality)).filter(Boolean) as IContextualMenuItem[],
    [items, modalities]
  );
  const menuItems = useMemo(() => items.filter(({ key }) => !modalities.includes(key as ModalityType)), [
    items,
    modalities,
  ]);

  const handleRemoveModality = useCallback(
    (modality: ModalityType) => () => {
      // const templateId = modalityTemplates[modality].templateId;
      if (modalities.length > 1) {
        const updatedModalities = modalities.filter((item) => item !== modality);
        setModalities(updatedModalities);
        setSelectedKey(updatedModalities[0] as string);
        // onTemplateChange?.(templateId);
      }
    },
    [modalities, setModalities, setSelectedKey]
  );

  const handleItemClick = useCallback(
    (_, item?: IContextualMenuItem) => {
      if (item?.key) {
        setModalities((current) => [...current, item.key as ModalityType]);
        setSelectedKey(item.key);
      }
    },
    [setModalities]
  );

  const handleLinkClicked = useCallback((item?: PivotItem) => {
    if (item?.props.itemKey) {
      setSelectedKey(item?.props.itemKey);
    }
  }, []);

  const handleUpdateResponseTemplate = React.useCallback(
    (partialResponse: TemplateResponse) => {
      if (lgOption?.templateId) {
        const mergedResponse = mergeWith({}, response, partialResponse, (_, srcValue) => {
          if (Array.isArray(srcValue)) {
            return srcValue;
          }
        });
        const mappedResponse = `[Activity
  ${(Object.values(mergedResponse) as { kind: string; value: unknown }[])
    .map(({ kind, value }) => {
      if (!value || (Array.isArray(value) && !value.length)) {
        return;
      }

      if (Array.isArray(value) && ['Attachments', 'SuggestedActions'].includes(kind)) {
        return `${kind} = ${value.join(' | ')}`;
      }

      if (typeof value === 'string') {
        return `${kind} = ${value}`;
      }
    })
    .filter(Boolean)
    .join('\n\t')}
]`;
        setResponse(mergedResponse);
        onTemplateChange(lgOption.templateId, mappedResponse);
      }
    },
    [lgOption, response]
  );

  const handleAttachmentLayoutChange = useCallback(
    (layout: string) => {
      // handleUpdateResponseTemplate({ AttachmentLayout: { kind: 'AttachmentLayout', value: [layout] }, });
    },
    [handleUpdateResponseTemplate]
  );

  const handleInputHintChange = useCallback(
    (inputHint: string) => {
      // handleUpdateResponseTemplate({ InputHint: { kind: 'InputHint', value: [inputHint], } });
    },
    [handleUpdateResponseTemplate]
  );

  const addMenuProps = React.useMemo<IContextualMenuProps>(
    () => ({
      items: menuItems,
      onItemClick: handleItemClick,
    }),
    [menuItems, handleItemClick]
  );

  return (
    <Stack>
      <Stack horizontal verticalAlign="center">
        <Pivot headersOnly selectedKey={selectedKey} styles={styles.tabs} onLinkClick={handleLinkClicked}>
          {pivotItems.map(({ key, text }) => (
            <PivotItem key={key} headerText={text} itemKey={key} />
          ))}
        </Pivot>
        {menuItems.filter((item) => item.itemType !== ContextualMenuItemType.Header).length && (
          <IconButton iconProps={addButtonIconProps} menuProps={addMenuProps} onRenderMenuIcon={() => null} />
        )}
      </Stack>

      <div ref={containerRef}>
        {renderModalityEditor(
          response,
          selectedKey,
          handleRemoveModality,
          onTemplateChange,
          handleAttachmentLayoutChange,
          handleInputHintChange,
          handleUpdateResponseTemplate,
          modalities.length === 1,
          lgOption,
          lgTemplates,
          memoryVariables
        )}
      </div>
    </Stack>
  );
});

export { ModalityPivot };
