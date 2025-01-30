---
date: 2024-11-15
authors: [2o1oo]
description: VBA Monologger documentation site just landed
categories:
  - Blog
links:
  - setup/setting-up-a-blog.md
  - plugins/blog.md
---

# VBA Monologger documentation site just landed

**Welcome! You're looking at the new website documentation, built with [Material for Mkdocs](https://squidfunk.github.io/mkdocs-material/), and deploy on [GitHub Pages](https://docs.github.com/en/pages).** Thanks to them for this wedding.

![img.png](welcome.png)


<!-- more -->

## What is GitHub Pages?

GitHub Pages is a service provided by GitHub that allows you to host static websites directly from your GitHub repositories. It’s particularly useful for developers, teams, and open-source projects, enabling them to create personal portfolios, project documentation, or organizational sites with minimal setup. The hosted websites are directly tied to the content of a repository, which means that updates to the repository automatically update the site.

For more information, visit the [GitHub Pages documentation](https://docs.github.com/en/pages).


## What is MkDocs and Material for MkDocs?

**MkDocs** is an open-source static site generator specifically designed for creating documentation websites. Written in Python, it offers a clean, developer-friendly experience for building sites quickly. Its focus on simplicity makes it a popular choice for technical documentation.

**Material for MkDocs** is a theme for MkDocs that enhances its capabilities with a modern and professional look. Inspired by Google’s Material Design principles, this theme provides a sleek, responsive design and a wealth of customization options, making your documentation more engaging and accessible.

Together, they create a powerful framework for writing and publishing technical documentation easily and efficiently. For more information, visit the [Material for MkDocs documentation](https://squidfunk.github.io/mkdocs-material/).


## Curious to see and understand how this site is built?

For those curious about how the VBA Monologger documentation site is constructed (in JAMStack we should say generated), the key lies in the structure of the `/docs` directory within the Git repository. This folder contains all the Markdown files that make up the content of the site, as well as other resources like images or custom scripts.

And the crucial piece of the puzzle is the `mkdocs.yml` configuration file located at the root of the repository. This file acts as the blueprint for the site, defining how pages are organized, specifying themes, and enabling features like search functionality or plugins.

Happy exploring, and have fun customizing your own documentation projects!

