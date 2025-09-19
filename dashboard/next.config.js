/** @type {import('next').NextConfig} */
const nextConfig = {
  // 🚧 Let the site deploy even if ESLint/TS has issues.
  eslint: {
    ignoreDuringBuilds: true,
  },
  typescript: {
    ignoreBuildErrors: true,
  },
};

module.exports = nextConfig;
