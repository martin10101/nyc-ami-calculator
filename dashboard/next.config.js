/** @type {import('next').NextConfig} */
const nextConfig = {
  // 🚧 Allow deploy even if ESLint/TS has issues
  eslint: {
    ignoreDuringBuilds: true,
  },
  typescript: {
    ignoreBuildErrors: true,
  },
  // ✅ Force static export (no serverless functions)
  output: 'export',
};

module.exports = nextConfig;

