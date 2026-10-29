function setGithubPat() {
  const pat = 'YOUR_GITHUB_PAT_HERE';
  if (!pat || pat.startsWith('YOUR_')) {
    throw new Error('Replace the placeholder token in Token.js with a GitHub Personal Access Token before running setGithubPat.');
  }
  PropertiesService.getScriptProperties().setProperty('GITHUB_PAT', pat);
  Logger.log('GITHUB_PAT saved to Script Properties.');
}
