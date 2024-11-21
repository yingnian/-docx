<template>
  <div id="app">
    <pre>
      {{ list }}
    </pre>
    <span style="background-color: blue; color: #fff; padding: 5px; cursor: pointer" @click="exportDoc">导出word</span>
  </div>
</template>

<script>
import { exportDocx } from './utils/doc.js'
export default {
  name: 'App',
  data() {
    return {
      list: {
        num: 12,
        dayNum: 20,
        options: [
          { index: 1, name: '测试数据', time: '2023-12-01', status: '这个是状态' },
          { index: 2, name: '测试数据1', time: '2023-12-01', status: '这个是状态' },
          { index: 3, name: '测试数据2', time: '2023-12-01', status: '这个是状态' }
        ],
        otherList: [
          { index: 1, desc: '描述信息1' },
          { index: 2, desc: '描述信息2' },
          { index: 3, desc: '描述信息3' }
        ],
        imgList: [
          // 这是一个 error 的 img
          { img: 'http://gips0.baidu.com/it/u=360272273692,151222483864&fm=3028&app=3028&f=JPEG&fmt=auto?w=960&h=1282220' },
          // 这是一个正常的 img
          { img: 'http://gips2.baidu.com/it/u=195724436,3554684702&fm=3028&app=3028&f=JPEG&fmt=auto?w=1280&h=960' }
        ]
      }
    }
  },
  methods: {
    exportDoc() {
      exportDocx('/template.docx', this.list, `ceshi.docx`)
    }
  }
}
</script>

<style>
#app {
  font-family: Avenir, Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  color: #2c3e50;
  margin-top: 60px;
}
</style>
