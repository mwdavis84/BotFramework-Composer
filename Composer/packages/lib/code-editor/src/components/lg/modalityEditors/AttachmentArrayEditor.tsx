// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { LgTemplate } from '@bfc/shared';
import formatMessage from 'format-message';
import { CommandButton, IButtonStyles } from 'office-ui-fabric-react/lib/Button';
import { FluentTheme } from '@uifabric/fluent-theme';
import {
  IContextualMenuItem,
  IContextualMenuProps,
  IContextualMenuItemProps,
} from 'office-ui-fabric-react/lib/ContextualMenu';
import React from 'react';

import { LGOption } from '../../../utils';
import { jsLgToolbarMenuClassName } from '../constants';

import { StringArrayItem } from './StringArrayItem';

const styles: { button: IButtonStyles } = {
  button: {
    root: {
      color: FluentTheme.palette.themePrimary,
      fontSize: FluentTheme.fonts.small.fontSize,
    },
  },
};

const addButtonMenuItemProps: Partial<IContextualMenuItemProps> = { styles: { label: { ...FluentTheme.fonts.small } } };

type AttachmentArrayEditorProps = {
  items: string[];
  selectedKey: string;
  lgTemplates?: readonly LgTemplate[];
  memoryVariables?: readonly string[];
  lgOption?: LGOption;
  onChange: (items: string[]) => void;
  onTemplateChange: (templateId: string, body?: string) => void;
};

const AttachmentArrayEditor = React.memo(
  ({ items, lgOption, lgTemplates, memoryVariables, onChange, onTemplateChange }: AttachmentArrayEditorProps) => {
    const containerRef = React.useRef<HTMLDivElement | null>(null);
    const [currentIndex, setCurrentIndex] = React.useState<number | null>(null);

    const handleChange = React.useCallback(
      (templateId: string) => (body?: string) => {
        onTemplateChange(templateId, body);
      },
      [items, onTemplateChange]
    );

    const handleFocus = React.useCallback(
      (index: number) => () => {
        setCurrentIndex(index);
      },
      [setCurrentIndex]
    );

    const handleRemove = React.useCallback(
      (index: number) => () => {
        const newItems = items.slice();
        newItems.splice(index, 1);
        onChange(newItems);
      },
      [items, onChange]
    );

    const handleAddTemplateClick = React.useCallback(
      (_, item?: IContextualMenuItem) => {
        const templateId = `HeroCard`;
        onChange([...items, templateId]);
        onTemplateChange(templateId, '[Test\n\ttext = 123\n]');
      },
      [items, onChange, onTemplateChange]
    );

    const newButtonMenuItems = React.useMemo<IContextualMenuItem[]>(
      () => [
        {
          key: 'addCustom',
          text: formatMessage('Add Custom'),
          itemProps: addButtonMenuItemProps,
        },
        {
          key: 'template',
          text: formatMessage('Create from templates'),
          itemProps: addButtonMenuItemProps,
          subMenuProps: {
            items: [
              {
                key: 'hero',
                text: formatMessage('Hero card'),
                onClick: handleAddTemplateClick,
                itemProps: addButtonMenuItemProps,
              },
              {
                key: 'thumbnail',
                text: formatMessage('Thumbnail card'),
                itemProps: addButtonMenuItemProps,
              },
              {
                key: 'signin',
                text: formatMessage('Sign-in card'),
                itemProps: addButtonMenuItemProps,
              },
              {
                key: 'animation',
                text: formatMessage('Animation card'),
                itemProps: addButtonMenuItemProps,
              },
              {
                key: 'video',
                text: formatMessage('Video card'),
                itemProps: addButtonMenuItemProps,
              },
              {
                key: 'audio',
                text: formatMessage('Audio card'),
                itemProps: addButtonMenuItemProps,
              },
              {
                key: 'adaptive',
                text: formatMessage('Adaptive card'),
                itemProps: addButtonMenuItemProps,
              },
              {
                key: 'url',
                text: formatMessage('Url'),
                itemProps: addButtonMenuItemProps,
              },
            ],
          },
        },
      ],
      [handleAddTemplateClick]
    );

    const addButtonMenuProps = React.useMemo<IContextualMenuProps>(() => ({ items: newButtonMenuItems }), [
      newButtonMenuItems,
    ]);

    React.useEffect(() => {
      const keydownHandler = (e: KeyboardEvent) => {
        if (e.key === 'Escape') {
          setCurrentIndex(null);
          // Remove empty variations
          onChange(items.filter(Boolean));
        }
      };

      const focusHandler = (e: FocusEvent) => {
        if (containerRef.current?.contains(e.target as Node)) {
          return;
        }

        if (
          !e
            .composedPath()
            .filter((n) => n instanceof Element)
            .map((n) => (n as Element).className)
            .some((c) => c.indexOf(jsLgToolbarMenuClassName) !== -1)
        ) {
          setCurrentIndex(null);
          // Remove empty variations
          onChange(items.filter(Boolean));
        }
      };

      document.addEventListener('keydown', keydownHandler);
      document.addEventListener('focusin', focusHandler);

      return () => {
        document.removeEventListener('keydown', keydownHandler);
        document.removeEventListener('focusin', focusHandler);
      };
    }, [items, onChange]);

    const templates = React.useMemo(() => {
      return items.map((name) => {
        return lgTemplates?.find((template) => template.name === name) || { name, body: '' };
      }, []);
    }, [items, lgTemplates]);

    return (
      <div ref={containerRef}>
        {templates.map(({ name, body }, key) => (
          <StringArrayItem
            key={key}
            editorMode="editor"
            lgOption={lgOption}
            lgTemplates={lgTemplates}
            memoryVariables={memoryVariables}
            mode={key === currentIndex ? 'edit' : 'view'}
            value={body}
            onFocus={handleFocus(key)}
            onLgChange={handleChange(name)}
            onRemove={handleRemove(key)}
          />
        ))}
        {currentIndex === null && (
          <CommandButton menuProps={addButtonMenuProps} styles={styles.button} onRenderMenuIcon={() => null}>
            {formatMessage('Add new attachment')}
          </CommandButton>
        )}
      </div>
    );
  }
);

export { AttachmentArrayEditor };
