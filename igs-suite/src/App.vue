<template lang="pug">
  #app
    one-slidebar
</template>

<script>
import oneSlidebar from "./components/ui/slidebar.vue"

export default {
  name: 'app',
  data () {
    return {
      msg: 'Welcome to Your Vue.js App'
    }
  },
  components: {
    oneSlidebar
  }
}
</script>

<style lang="scss">
body{
  margin: 0;
  position: absolute;
  width: 100%;
  min-height: 100%;

}
#app {
  position: absolute;
  width: 100%;
  min-height: 100%;
  display: grid;
  grid-template-columns: 48px 1fr;
  grid-template-rows: 48px 1fr;
}

h1, h2 {
  font-weight: normal;
}

ul {
  list-style-type: none;
  padding: 0;
}

li {
  display: inline-block;
  margin: 0 10px;
}

a {
  color: #42b983;
}
</style>
