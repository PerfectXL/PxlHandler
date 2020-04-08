<template>
  <div id="app" class="container mt-5">
    <div class="row">
      <div class="col">
        <div class="jumbotron border">
          <h1 class="display-5">PXL Link Demo</h1>
          <form>
            <div class="form-group">
              <label for="filename">File name </label>
              <input class="form-control" v-model="filename" />
              <small class="form-text text-muted">Including path if file is not open in Excel</small>
            </div>
            <div class="form-group">
              <label for="worksheetName">Worksheet name</label>
              <input class="form-control" v-model="worksheetName" />
            </div>
            <div class="form-group">
              <label for="rangeAddress">Cell (or range) address</label>
              <input class="form-control" v-model="rangeAddress" />
              <small class="form-text text-muted">May be left empty</small>
            </div>
            <hr />
            <p class="float-right">
              <a class="btn btn-primary" :class="{disabled: !formIsValid}" :href="formIsValid ? link : ''">{{ link }}</a>
            </p>
          </form>
        </div>
      </div>
    </div>
  </div>
</template>

<script>
export default {
  name: "App",
  data: function() {
    return {
      filename: "",
      worksheetName: "",
      rangeAddress: "",
    };
  },
  computed: {
    formIsValid() {
      return this.filename && this.worksheetName;
    },
    link() {
      return 'pxl://jump-to-cell/' + [this.filename, this.worksheetName, this.rangeAddress].filter(s => s.length > 0).map(encodeURIComponent).join('/');
    },
  },
};
</script>
