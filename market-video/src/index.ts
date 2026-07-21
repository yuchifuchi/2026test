import { registerRoot } from 'remotion';
import React from 'react';
import { Composition } from 'remotion';
import { Short } from './Short';
import { Long } from './Long';
import brief from '../data/brief.json';
import longform from '../data/longform.json';

const FPS = 30;
const shortSec = Math.max(16, Math.ceil(((brief as any).narrationDuration ?? 0) + 1.5));
const longSec = Math.max(80, Math.ceil(((longform as any).narrationDuration ?? 0) + 3));

const Root: React.FC = () => {
  return React.createElement(
    React.Fragment,
    null,
    React.createElement(Composition, { id: 'Short', component: Short, durationInFrames: FPS * shortSec, fps: FPS, width: 1080, height: 1920 }),
    React.createElement(Composition, { id: 'Long', component: Long, durationInFrames: FPS * longSec, fps: FPS, width: 1920, height: 1080 }),
  );
};

registerRoot(Root);
