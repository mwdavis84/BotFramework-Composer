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
import { structuredResponseToString } from '../../utils/structuredResponse';

import { AttachmentModalityEditor } from './modalityEditors/AttachmentModalityEditor';
import { SpeechModalityEditor } from './modalityEditors/SpeechModalityEditor';
import { SuggestedActionsModalityEditor } from './modalityEditors/SuggestedActionsModalityEditor';
import { TextModalityEditor } from './modalityEditors/TextModalityEditor';
import {
  AttachmentsStructuredResponseItem,
  SpeechStructuredResponseItem,
  SuggestedActionsStructuredResponseItem,
  TextStructuredResponseItem,
  ModalityType,
  PartialStructuredResponse,
  AttachmentLayoutStructuredResponseItem,
  InputHintStructuredResponseItem,
  modalityType,
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

const renderModalityEditor = ({
  modality,
  removeModalityDisabled,
  structuredResponse,
  lgOption,
  lgTemplates,
  memoryVariables,
  onRemoveModality,
  onTemplateChange,
  onAttachmentLayoutChange,
  onInputHintChange,
  onUpdateResponseTemplate,
}: {
  modality: string;
  removeModalityDisabled: boolean;
  structuredResponse?: PartialStructuredResponse;
  lgOption?: LGOption;
  lgTemplates?: readonly LgTemplate[];
  memoryVariables?: readonly string[];
  onRemoveModality: (modality: ModalityType) => void;
  onTemplateChange: (templateId: string, body?: string) => void;
  onAttachmentLayoutChange: (layout: string) => void;
  onInputHintChange: (inputHintString: string) => void;
  onUpdateResponseTemplate: (response: PartialStructuredResponse) => void;
}) => {
  const commonProps = {
    lgOption,
    lgTemplates,
    memoryVariables,
    removeModalityDisabled,
    onTemplateChange,
    onUpdateResponseTemplate,
    onRemoveModality,
  };

  switch (modality) {
    case 'Attachments':
      return (
        <AttachmentModalityEditor
          {...commonProps}
          attachmentLayout={(structuredResponse?.AttachmentLayout as AttachmentLayoutStructuredResponseItem)?.value}
          response={structuredResponse?.Attachments as AttachmentsStructuredResponseItem}
          onAttachmentLayoutChange={onAttachmentLayoutChange}
        />
      );
    case 'Speak':
      return (
        <SpeechModalityEditor
          {...commonProps}
          inputHint={(structuredResponse?.InputHint as InputHintStructuredResponseItem)?.value}
          response={structuredResponse?.Speak as SpeechStructuredResponseItem}
          onInputHintChange={onInputHintChange}
        />
      );
    case 'SuggestedActions':
      return (
        <SuggestedActionsModalityEditor
          {...commonProps}
          response={structuredResponse?.SuggestedActions as SuggestedActionsStructuredResponseItem}
        />
      );
    case 'Text':
      return <TextModalityEditor {...commonProps} response={structuredResponse?.Text as TextStructuredResponseItem} />;
  }
};

const getInitialModalities = (structuredResponse?: PartialStructuredResponse): ModalityType[] => {
  const modalities = Object.keys(structuredResponse || {}).filter((m) =>
    modalityType.includes(m as ModalityType)
  ) as ModalityType[];
  return modalities.length ? modalities : ['Text'];
};

type Props = {
  lgOption?: LGOption;
  lgTemplates?: readonly LgTemplate[];
  memoryVariables?: readonly string[];
  structuredResponse?: PartialStructuredResponse;
  onTemplateChange?: (templateId: string, body?: string) => void;
};

export const ModalityPivot = React.memo((props: Props) => {
  const {
    lgOption,
    lgTemplates,
    memoryVariables,
    structuredResponse: initialStructuredResponse,
    onTemplateChange = () => {},
  } = props;

  const [structuredResponse, setStructuredResponse] = React.useState(initialStructuredResponse);
  const [modalities, setModalities] = useState<ModalityType[]>(getInitialModalities(structuredResponse));
  const [selectedModality, setSelectedModality] = useState<string>(modalities[0] as string);

  const containerRef = useRef<HTMLDivElement>(null);

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

  const onRemoveModality = useCallback(
    (modality: ModalityType) => {
      // const templateId = modalityTemplates[modality].templateId;
      if (modalities.length > 1) {
        const updatedModalities = modalities.filter((item) => item !== modality);
        setModalities(updatedModalities);
        setSelectedModality(updatedModalities[0] as string);

        if (lgOption?.templateId) {
          const mergedResponse = mergeWith({}, structuredResponse) as PartialStructuredResponse;
          delete mergedResponse[modality];

          setStructuredResponse(mergedResponse);
          const mappedResponse = structuredResponseToString(mergedResponse);
          onTemplateChange(lgOption.templateId, mappedResponse);
        }
      }
    },
    [modalities, setModalities, setSelectedModality, lgOption]
  );

  const onModalityAddMenuItemClick = useCallback(
    (_, item?: IContextualMenuItem) => {
      if (item?.key) {
        setModalities((current) => [...current, item.key as ModalityType]);
        setSelectedModality(item.key);
      }
    },
    [setModalities]
  );

  const onPivotChange = useCallback((item?: PivotItem) => {
    if (item?.props.itemKey) {
      setSelectedModality(item?.props.itemKey);
    }
  }, []);

  const onUpdateResponseTemplate = React.useCallback(
    (partialResponse: PartialStructuredResponse) => {
      if (lgOption?.templateId) {
        const mergedResponse = mergeWith({}, structuredResponse, partialResponse, (_, srcValue) => {
          if (Array.isArray(srcValue)) {
            return srcValue;
          }
        });
        setStructuredResponse(mergedResponse);

        const mappedResponse = structuredResponseToString(mergedResponse);

        onTemplateChange(lgOption.templateId, mappedResponse);
      }
    },
    [lgOption, structuredResponse]
  );

  const onAttachmentLayoutChange = useCallback(
    (layout: string) => {
      onUpdateResponseTemplate({
        AttachmentLayout: { kind: 'AttachmentLayout', value: layout } as AttachmentLayoutStructuredResponseItem,
      });
    },
    [onUpdateResponseTemplate]
  );

  const onInputHintChange = useCallback(
    (inputHint: string) => {
      onUpdateResponseTemplate({
        InputHint:
          inputHint !== 'none'
            ? ({ kind: 'InputHint', value: inputHint } as InputHintStructuredResponseItem)
            : undefined,
      });
    },
    [onUpdateResponseTemplate]
  );

  const addMenuProps = React.useMemo<IContextualMenuProps>(
    () => ({
      items: menuItems,
      onItemClick: onModalityAddMenuItemClick,
    }),
    [menuItems, onModalityAddMenuItemClick]
  );

  return (
    <Stack>
      <Stack horizontal verticalAlign="center">
        <Pivot headersOnly selectedKey={selectedModality} styles={styles.tabs} onLinkClick={onPivotChange}>
          {pivotItems.map(({ key, text }) => (
            <PivotItem key={key} headerText={text} itemKey={key} />
          ))}
        </Pivot>
        {menuItems.filter((item) => item.itemType !== ContextualMenuItemType.Header).length && (
          <IconButton iconProps={addButtonIconProps} menuProps={addMenuProps} onRenderMenuIcon={() => null} />
        )}
      </Stack>

      <div ref={containerRef}>
        {renderModalityEditor({
          removeModalityDisabled: modalities.length === 1,
          structuredResponse,
          modality: selectedModality,
          lgOption,
          lgTemplates,
          memoryVariables,
          onRemoveModality,
          onTemplateChange,
          onAttachmentLayoutChange,
          onInputHintChange,
          onUpdateResponseTemplate,
        })}
      </div>
    </Stack>
  );
});
