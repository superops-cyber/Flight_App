# Simple Dev Workflow (/dev URL)

Use this for normal development.

1. Open your Apps Script web app **/dev** URL in the browser.
2. Make code changes locally.
3. Push changes:

```bash
npm run dev:push
```

4. Hard refresh browser:

```text
Cmd+Shift+R
```

5. Test. Repeat.

## Helpful Commands

- Push + reminder message:

```bash
npm run dev:loop
```

- Open web app page from terminal:

```bash
npm run dev:webapp
```

## Release Commands (only when publishing /exec)

- List deployments:

```bash
npm run release:list
```

- Create a new version:

```bash
npm run release:version
```

Legacy release script is still available:

```bash
npm run release:deploy:legacy
```
