<script setup lang="ts"></script>

<template>
  <div class="app-container">
    <!-- 动态背景层 -->
    <div class="background-layer">
      <!-- 渐变基底 -->
      <div class="gradient-base"></div>

      <!-- 浮动几何元素 -->
      <div class="floating-shape shape-1"></div>
      <div class="floating-shape shape-2"></div>
      <div class="floating-shape shape-3"></div>
      <div class="floating-shape shape-4"></div>
      <div class="floating-shape shape-5"></div>

      <!-- 网格纹理 -->
      <div class="grid-overlay"></div>
    </div>

    <!-- 内容层 -->
    <div class="content-layer">
      <router-view />
    </div>
  </div>
</template>

<style scoped>
.app-container {
  position: relative;
  width: 100%;
  height: 100vh;
  overflow: hidden;
}

/* ============ 背景层 ============ */
.background-layer {
  position: fixed;
  top: 0;
  left: 0;
  width: 100%;
  height: 100%;
  z-index: 0;
  overflow: hidden;
}

/* 渐变基底 */
.gradient-base {
  position: absolute;
  width: 100%;
  height: 100%;
  background: linear-gradient(
    135deg,
    #f0f4ff 0%,
    #e8f0fe 25%,
    #f3f0ff 50%,
    #f8f4ff 75%,
    #fef7f0 100%
  );
  animation: gradientShift 40s ease-in-out infinite alternate;
}

@keyframes gradientShift {
  0% {
    filter: hue-rotate(0deg) brightness(1);
  }
  50% {
    filter: hue-rotate(5deg) brightness(1.02);
  }
  100% {
    filter: hue-rotate(0deg) brightness(1);
  }
}

/* 浮动几何图形 */
.floating-shape {
  position: absolute;
  border-radius: 50%;
  filter: blur(60px);
  opacity: 0.3;
  will-change: transform;
}

.shape-1 {
  width: 500px;
  height: 500px;
  background: radial-gradient(circle, rgba(99, 102, 241, 0.4) 0%, transparent 70%);
  top: -10%;
  left: -10%;
  animation: float1 35s ease-in-out infinite;
}

.shape-2 {
  width: 400px;
  height: 400px;
  background: radial-gradient(circle, rgba(147, 51, 234, 0.35) 0%, transparent 70%);
  top: 60%;
  right: -5%;
  animation: float2 40s ease-in-out infinite;
}

.shape-3 {
  width: 350px;
  height: 350px;
  background: radial-gradient(circle, rgba(59, 130, 246, 0.3) 0%, transparent 70%);
  bottom: -10%;
  left: 30%;
  animation: float3 30s ease-in-out infinite;
}

.shape-4 {
  width: 450px;
  height: 450px;
  background: radial-gradient(circle, rgba(168, 85, 247, 0.25) 0%, transparent 70%);
  top: 40%;
  left: 50%;
  animation: float4 45s ease-in-out infinite;
}

.shape-5 {
  width: 300px;
  height: 300px;
  background: radial-gradient(circle, rgba(99, 102, 241, 0.28) 0%, transparent 70%);
  top: 20%;
  right: 25%;
  animation: float5 38s ease-in-out infinite;
}

/* 浮动动画 */
@keyframes float1 {
  0%,
  100% {
    transform: translate(0, 0) scale(1);
  }
  33% {
    transform: translate(60px, 80px) scale(1.1);
  }
  66% {
    transform: translate(-40px, 120px) scale(0.95);
  }
}

@keyframes float2 {
  0%,
  100% {
    transform: translate(0, 0) scale(1);
  }
  33% {
    transform: translate(-80px, -60px) scale(1.05);
  }
  66% {
    transform: translate(50px, -100px) scale(0.98);
  }
}

@keyframes float3 {
  0%,
  100% {
    transform: translate(0, 0) rotate(0deg) scale(1);
  }
  50% {
    transform: translate(100px, -80px) rotate(10deg) scale(1.08);
  }
}

@keyframes float4 {
  0%,
  100% {
    transform: translate(0, 0) scale(1);
  }
  25% {
    transform: translate(-70px, 90px) scale(0.92);
  }
  50% {
    transform: translate(50px, -50px) scale(1.1);
  }
  75% {
    transform: translate(80px, 70px) scale(0.95);
  }
}

@keyframes float5 {
  0%,
  100% {
    transform: translate(0, 0) rotate(0deg);
  }
  50% {
    transform: translate(-90px, 110px) rotate(-15deg);
  }
}

/* 网格纹理 */
.grid-overlay {
  position: absolute;
  width: 100%;
  height: 100%;
  background-image:
    linear-gradient(rgba(99, 102, 241, 0.03) 1px, transparent 1px),
    linear-gradient(90deg, rgba(99, 102, 241, 0.03) 1px, transparent 1px);
  background-size: 50px 50px;
  opacity: 0.6;
  animation: gridPulse 20s ease-in-out infinite;
}

@keyframes gridPulse {
  0%,
  100% {
    opacity: 0.5;
  }
  50% {
    opacity: 0.7;
  }
}

/* ============ 内容层 ============ */
.content-layer {
  position: relative;
  width: 100%;
  height: 100%;
  z-index: 1;
  overflow: auto;
}

/* 性能优化 */
@media (prefers-reduced-motion: reduce) {
  .floating-shape,
  .gradient-base,
  .grid-overlay {
    animation: none;
  }
}
</style>
