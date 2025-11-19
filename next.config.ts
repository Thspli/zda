import type { NextConfig } from 'next'

const nextConfig: NextConfig = {
  webpack: (config: any, { isServer }: { isServer: boolean }) => {
    if (isServer) {
      // Não tenta bundlar módulos nativos do Node
      config.externals = config.externals || [];
      config.externals.push('child_process');
    }
    return config;
  },
  experimental: {
    serverActions: {
      bodySizeLimit: '50mb',
    },
  },
}

export default nextConfig