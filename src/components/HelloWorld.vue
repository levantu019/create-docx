<template>
  <div class="hello">
    <button @click="all">Download All</button>
    <button @click="child">Download Childs</button>
  </div>
</template>

<script>
import { saveAs } from "file-saver";
import { Packer } from "docx";
import {
  DocumentAllEquipments,
  DocumentChildEquipments,
} from "../docx_functions";
import { parent_childs, all_equipment } from "../docx_data";

export default {
  name: "HelloWorld",
  props: {
    msg: String,
  },
  methods: {
    all() {
      Packer.toBlob(DocumentAllEquipments(all_equipment)).then((blob) => {
        saveAs(blob, "example.docx");
      });
    },
    child() {
      Packer.toBlob(DocumentChildEquipments(parent_childs)).then((blob) => {
        saveAs(blob, "example.docx");
      });
    },
  },
};
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style scoped>
</style>
