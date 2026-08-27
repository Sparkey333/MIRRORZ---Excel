#!/usr/bin/env node
// Thin launcher so the package can be run straight from source during
// development and from the built output once compiled.
import { main } from '../src/index.ts';
process.exit(await main(process.argv));
