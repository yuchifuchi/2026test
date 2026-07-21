import { registerRoot } from 'remotion';
import React from 'react';
import { Composition } from 'remotion';
import { Short } from './Short';
import brief from '../data/brief.json';

const FPS = 30;
const narrSec = (brief as any).narrationDuration ?? 0;
const durationSec = Math.max(16, Math.ceil(narrSec + 1.5));

const Root: React.FC = () => {
  return React.createElement(Composition, {
    id: 'Short',
    component: Short,
    durationInFrames: FPS * durationSec,
    fps: FPS,
    width: 1080,
    height: 1920,
  });
};

registerRoot(Root);
