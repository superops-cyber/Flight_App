const { execSync } = require('child_process');

function main() {
  const output = execSync('clasp deployments', { encoding: 'utf8' });
  const headMatch = output.match(/^\s*-\s+([^\s]+)\s+@HEAD/m);

  if (!headMatch) {
    throw new Error('Could not find the current @HEAD deployment ID.');
  }

  const deploymentId = headMatch[1];
  console.log('Opening HEAD deployment:', deploymentId);
  execSync(`clasp open-web-app ${deploymentId}`, { stdio: 'inherit' });
}

main();