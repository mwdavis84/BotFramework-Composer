// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

export const activityTemplateType = 'Activity';
export const emptyTemplateBodyRegex = /^$|-(\s)?/;

export const jsLgToolbarMenuClassName = 'js-lg-toolbar-menu';

type AttachmentCard = 'hero' | 'thumbnail' | 'signin' | 'animation' | 'video' | 'audio'; // | 'adaptive' | 'url';

export const cardTemplates: Record<AttachmentCard, string> = {
  hero: `[HeroCard
  title =
  subtitle =
  text =
  images =
  buttons =
]
`,
  thumbnail: `[ThumbnailCard
  title =
  subtitle =
  text =
  image =
  buttons =
]`,
  signin: `[SigninCard
    text =
    buttons =
]`,
  animation: `[AnimationCard
    title =
    subtitle =
    image =
    media =
]`,
  video: `[VideoCard
    title =
    subtitle =
    text =
    image =
    media =
    buttons =
]`,
  audio: `[AudioCard
    title =
    subtitle =
    text =
    image =
    media =
    buttons =
]`,
};
