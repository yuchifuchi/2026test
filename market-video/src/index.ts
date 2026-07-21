import { registerRoot } from 'remotion';
import React from 'react';
import { Composition } from 'remotion';
import { Short } from './Short';

const Root: React.FC = () => {
  return React.createElement(Composition, {
    id: 'Short',
    component: Short,
    durationInFrames: 30 * 16,
    fps: 30,
    width: 1080,
    height: 1920,
  });
};

registerRoot(Root);
