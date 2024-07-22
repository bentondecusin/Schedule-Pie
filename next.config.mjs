/** @type {import('next').NextConfig} */
const nextConfig = {
  serverRuntimeConfig: {
    NEXT_PUBLIC_CLIENT_ID: process.env.NEXT_PUBLIC_CLIENT_ID,
  },
};

export default nextConfig;
